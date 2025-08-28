package main

import (
	"flag"
	"fmt"
)


func main() {
	var verbose bool 
	flag.BoolVar(&verbose, "v", false, "Enable verbose mode")
    flag.BoolVar(&verbose, "verbose", false, "Enable verbose mode")
	flag.Parse()
	var args = flag.Args()
	if len(args) < 1 {
		fmt.Println("Please provide the target file path")
		return
	}
	var targetFile = args[0]
	fmt.Println("Target file:", targetFile )
	if verbose {
		fmt.Println("Verbose mode is on")
	}
}