// Parts Analysis launcher: serves the embedded web/ folder on localhost and opens the browser.
// Build for Windows:  GOOS=windows GOARCH=amd64 go build -trimpath -ldflags "-s -w" -o "Parts Analysis.exe" .
package main

import (
	"embed"
	"fmt"
	"io/fs"
	"net"
	"net/http"
	"os"
	"os/exec"
	"runtime"
)

//go:embed web
var webFiles embed.FS

// 4173 first: browser data (localStorage) is tied to the address, so keep it stable across versions.
var ports = []int{4173, 4174, 4175, 4180, 8099, 8123}

func main() {
	site, err := fs.Sub(webFiles, "web")
	if err != nil {
		fail(err)
	}

	var ln net.Listener
	var port int
	for _, p := range ports {
		if ln, err = net.Listen("tcp", fmt.Sprintf("127.0.0.1:%d", p)); err == nil {
			port = p
			break
		}
	}
	if ln == nil {
		fail(fmt.Errorf("no free port among %v", ports))
	}

	url := fmt.Sprintf("http://localhost:%d/", port)
	fmt.Println()
	fmt.Println("  Parts Analysis is running at", url)
	fmt.Println("  Keep this window open while you use it. Close it to stop.")
	fmt.Println("  Your data is saved in the browser, so closing this window does not lose it.")
	fmt.Println()
	openBrowser(url)

	handler := http.FileServer(http.FS(site))
	if err := http.Serve(ln, http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Cache-Control", "no-cache")
		handler.ServeHTTP(w, r)
	})); err != nil {
		fail(err)
	}
}

func openBrowser(url string) {
	var cmd *exec.Cmd
	switch runtime.GOOS {
	case "windows":
		cmd = exec.Command("rundll32", "url.dll,FileProtocolHandler", url)
	case "darwin":
		cmd = exec.Command("open", url)
	default:
		cmd = exec.Command("xdg-open", url)
	}
	if err := cmd.Start(); err != nil {
		fmt.Println("  Could not open the browser automatically. Open this address yourself:", url)
	}
}

func fail(err error) {
	fmt.Println("  Could not start the local server:", err)
	fmt.Print("  Press Enter to close...")
	fmt.Scanln()
	os.Exit(1)
}
