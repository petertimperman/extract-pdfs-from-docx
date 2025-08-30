package main

import (
	"archive/zip"
	"encoding/xml"
	"flag"
	"fmt"
	"io"
	"strings"

	// "io"
	"log"
	"os"
)

// type VObject struct {

type EmbeddedObjectRefs struct {
	emfID string
	pdfID string
}

const ADOBE_TYPE_PROG_ID = "Acrobat.Document.DC"

func findEmbeddedObjects(documentXMLFile io.Reader) {
	documentDecoder := xml.NewDecoder(documentXMLFile)
	var inObjectTag bool
	var embeddedObjectRefs []EmbeddedObjectRefs
	var currentEmbeddedObject EmbeddedObjectRefs
	for {
		tok, err := documentDecoder.Token()
		if err != nil {
			break
		}
		switch element := tok.(type) {
		case xml.StartElement:
			localName := strings.TrimSpace(element.Name.Local) // Local names have whitespace sometimes=
			if localName == "object" {
				inObjectTag = true
				currentEmbeddedObject = EmbeddedObjectRefs{}
			}
			if inObjectTag {
				switch localName {
				case "OLEObject":
					tempPdfID := ""
					for _, attr := range element.Attr {
						if attr.Name.Local == "ProgID" && attr.Value == ADOBE_TYPE_PROG_ID {
							currentEmbeddedObject.pdfID = attr.Value
						} else if attr.Name.Local == "id" {
							currentEmbeddedObject.pdfID = attr.Value
						}
					}
					if tempPdfID != "" {
						currentEmbeddedObject.pdfID = tempPdfID
					}

				case "imagedata":
					for _, attr := range element.Attr {
						if attr.Name.Local == "id" {
							currentEmbeddedObject.emfID = attr.Value
						}
					}
				}
			}
		case xml.EndElement:
			localName := strings.TrimSpace(element.Name.Local)
			if localName == "object" {
				inObjectTag = false
				if currentEmbeddedObject.pdfID != "" && currentEmbeddedObject.emfID != "" {
					fmt.Printf("Found embedded PDF: %+v\n", currentEmbeddedObject)
					embeddedObjectRefs = append(embeddedObjectRefs, currentEmbeddedObject)
				}
			}
		}
	}
}


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
	if targetFile[len(targetFile)-5:] != ".docx" {
		log.Fatal("The target file must be a docx file")
		os.Exit(1)
	}
	fmt.Println("Target file:", targetFile)
	if verbose {
		fmt.Println("Verbose mode is on")
	}
	wordDirReader, err := zip.OpenReader(targetFile)
	if err != nil {
		log.Fatal(err)
	}
	defer wordDirReader.Close()
	documentFile, err := wordDirReader.Open("word/document.xml")
	if err != nil {
		log.Fatal(err)
	}
	defer documentFile.Close()
	log.Println("Successfully opened word/document.xml")
	findEmbeddedObjects(documentFile)	

}
