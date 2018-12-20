all:
	go build -tags dev -race

build-release:
	go build -tags release

build-no-cli:
	go build -tags release -ldflags -H=windowsgui -o bin.exe