package main

import (
	"flag"
	"fmt"
	"log"
	"net/http"
	"os"
	"os/exec"
	"path/filepath"
	"time"
)

func main() {
	port := flag.Int("port", 8080, "port to listen on")
	dir := flag.String("dir", "site", "path to the mkdocs site directory")
	host := flag.String("host", "127.0.0.1", "host address to bind to")
	noBrowser := flag.Bool("no-browser", false, "do not open the browser on startup")
	flag.Parse()

	// Resolve to absolute path for clarity in log output.
	absDir, err := filepath.Abs(*dir)
	if err != nil {
		log.Fatalf("Error resolving path: %v", err)
	}

	// Verify the directory exists.
	info, err := os.Stat(absDir)
	if err != nil {
		log.Fatalf("Cannot access site directory %q: %v", absDir, err)
	}
	if !info.IsDir() {
		log.Fatalf("%q is not a directory", absDir)
	}

	fs := http.FileServer(http.Dir(absDir))
	http.Handle("/", fs)

	addr := fmt.Sprintf("%s:%d", *host, *port)
	url := fmt.Sprintf("http://%s", addr)
	fmt.Printf("Serving MkDocs site from %s\n", absDir)
	fmt.Printf("Listening on %s\n", url)

	if !*noBrowser {
		go func() {
			time.Sleep(500 * time.Millisecond)
			exec.Command("rundll32", "url.dll,FileProtocolHandler", url).Start()
		}()
	}

	if err := http.ListenAndServe(addr, nil); err != nil {
		log.Fatalf("Server error: %v", err)
	}
}
