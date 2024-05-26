package main

import (
	"archive/zip"
	"bytes"
	"encoding/csv"
	"fmt"
	"io"
	"net/http"
	"os"
	"strconv"
	"strings"

	"github.com/gorilla/handlers"
	"github.com/xuri/excelize/v2"
)

// CleanSpreadsheet function to process the uploaded file
func CleanSpreadsheet(filePath string) (string, string, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return "", "", err
	}
	defer f.Close()

	var creditCSV, debitCSV strings.Builder
	creditWriter := csv.NewWriter(&creditCSV)
	defer creditWriter.Flush()
	debitWriter := csv.NewWriter(&debitCSV)
	defer debitWriter.Flush()

	for _, sheet := range f.GetSheetList() {
		// Remove merged cells
		mergedCells, err := f.GetMergeCells(sheet)
		if err != nil {
			return "", "", err
		}
		for _, mc := range mergedCells {
			err = f.UnmergeCell(sheet, mc.GetStartAxis(), mc.GetEndAxis())
			if err != nil {
				return "", "", err
			}
		}

		// Remove the first 25 rows
		for i := 1; i <= 25; i++ {
			err := f.RemoveRow(sheet, i)
			if err != nil {
				return "", "", err
			}
		}

		// Remove the last 14 rows
		rows, err := f.GetRows(sheet)
		if err != nil {
			return "", "", err
		}
		for i := len(rows) - 14; i < len(rows); i++ {
			err := f.RemoveRow(sheet, i+1)
			if err != nil {
				return "", "", err
			}
		}

		// Re-read rows after removals
		rows, err = f.GetRows(sheet)
		if err != nil {
			return "", "", err
		}

		if len(rows) == 0 {
			fmt.Printf("No rows found in sheet %s.\n", sheet)
			continue
		}

		for rowIndex, row := range rows {
			// Skip header row or rows without sufficient columns
			if rowIndex == 0 || len(row) < 39 {
				continue
			}

			amountStr := row[37]
			amountStr = strings.Replace(amountStr, ",", "", -1)

			// Handle empty or invalid amount strings
			if amountStr == "" || amountStr == "Amount" {
				continue
			}

			amount, err := strconv.ParseFloat(amountStr, 64)
			if err != nil {
				fmt.Println("Error parsing amount:", err)
				continue
			}

			formattedAmount := strconv.FormatFloat(amount, 'f', -1, 64)
			newRow := []string{row[0], row[24], formattedAmount}

			// Check if the amount is negative for credits
			if strings.HasPrefix(amountStr, "-") {
				// Convert the amount to positive
				positiveAmount := strconv.FormatFloat(-amount, 'f', -1, 64)
				newRow[2] = positiveAmount
				err = creditWriter.Write(newRow)
				if err != nil {
					return "", "", err
				}
			} else {
				err = debitWriter.Write(newRow)
				if err != nil {
					return "", "", err
				}
			}
		}

		creditWriter.Flush()
		debitWriter.Flush()
	}

	return creditCSV.String(), debitCSV.String(), nil
}

func uploadHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method != "POST" {
		http.Error(w, "Only POST method is allowed", http.StatusMethodNotAllowed)
		return
	}

	file, _, err := r.FormFile("file")
	if err != nil {
		http.Error(w, "Unable to read file from form", http.StatusBadRequest)
		return
	}
	defer file.Close()

	tmpFile, err := os.CreateTemp("", "uploaded-*.xlsx")
	if err != nil {
		http.Error(w, "Unable to create temporary file", http.StatusInternalServerError)
		return
	}
	defer os.Remove(tmpFile.Name())

	if _, err := io.Copy(tmpFile, file); err != nil {
		http.Error(w, "Unable to save uploaded file", http.StatusInternalServerError)
		return
	}

	creditCSV, debitCSV, err := CleanSpreadsheet(tmpFile.Name())
	if err != nil {
		http.Error(w, "Error processing file: "+err.Error(), http.StatusInternalServerError)
		return
	}

	if creditCSV == "" && debitCSV == "" {
		http.Error(w, "No data processed from the file", http.StatusInternalServerError)
		return
	}

	// Create a zip archive in memory
	buf := new(bytes.Buffer)
	zipWriter := zip.NewWriter(buf)

	// Add credits.csv to the zip archive
	creditFile, err := zipWriter.Create("credits.csv")
	if err != nil {
		http.Error(w, "Error creating zip file: "+err.Error(), http.StatusInternalServerError)
		return
	}
	creditFile.Write([]byte(creditCSV))

	// Add debits.csv to the zip archive
	debitFile, err := zipWriter.Create("debits.csv")
	if err != nil {
		http.Error(w, "Error creating zip file: "+err.Error(), http.StatusInternalServerError)
		return
	}
	debitFile.Write([]byte(debitCSV))

	// Close the zip archive
	if err := zipWriter.Close(); err != nil {
		http.Error(w, "Error closing zip file: "+err.Error(), http.StatusInternalServerError)
		return
	}

	// Set response headers
	w.Header().Set("Content-Type", "application/zip")
	w.Header().Set("Content-Disposition", "attachment; filename=processed_files.zip")

	// Write the zip archive to the response
	if _, err := w.Write(buf.Bytes()); err != nil {
		http.Error(w, "Error writing response: "+err.Error(), http.StatusInternalServerError)
		return
	}
}

func main() {
	// Create a new router
	router := http.NewServeMux()

	// Handle the upload route
	router.HandleFunc("/upload", uploadHandler)

	// Add CORS middleware
	corsHandler := handlers.CORS(
		handlers.AllowedOrigins([]string{"*"}),    // Allow requests from any origin
		handlers.AllowedMethods([]string{"POST"}), // Allow only POST requests
	)

	// Wrap the router with the CORS middleware
	http.ListenAndServe(":6666", corsHandler(router))
}
