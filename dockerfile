FROM --platform=linux/amd64 golang:1.22-alpine

WORKDIR /app

COPY go.mod go.sum ./

RUN go mod download

COPY . .

RUN GOARCH=arm64 go build -o main .

EXPOSE 6666
CMD ["/app/main"]
