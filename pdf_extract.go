package main

import (
	"archive/zip"
	"encoding/xml"
	"flag"
	"fmt"
	"io"
	"log"
	"os"
	"strings"
)

// type VObject struct {

type EmbeddedObjectRefs struct {
	emfID   string
	pdfID   string
}
type EmbeddedObjectPaths struct {
	emfPath string
	pdfPath string
}
const ADOBE_TYPE_PROG_ID = "Acrobat.Document.DC"

func findEmbeddedObjects(documentXMLFile io.Reader) []EmbeddedObjectRefs {
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
					embeddedObjectRefs = append(embeddedObjectRefs, currentEmbeddedObject)
				}
			}
		}
	}
	return embeddedObjectRefs
}

// func findTitleInEmf(emfFile io.Reader) string {

// }

func getRelToPaths(relXMLFile io.Reader) map[string]string {
	relToPathMap := make(map[string]string)
	relationshipDecoder := xml.NewDecoder(relXMLFile)

	for {
		token, err := relationshipDecoder.Token()
		if err != nil {
			break
		}
		switch element := token.(type) {
		case xml.StartElement:
			var id, target string
			for _, attr := range element.Attr {
				if attr.Name.Local == "Id" {
					id = attr.Value
				} else if attr.Name.Local == "Target" {
					if (strings.HasPrefix(attr.Value, "media/") || strings.HasPrefix(attr.Value, "embeddings/")) && (strings.HasSuffix(attr.Value, ".bin") || strings.HasSuffix(attr.Value, ".emf")) {
						target = attr.Value
					}
				}
			}
			if id != "" && target != "" {
				relToPathMap[id] = target
			}
		}
	}
	return relToPathMap
}

func getFileNameFromEmfPath(emfPath string) string {
	emfFile , err := os.Open(emfPath)
	if err != nil {
		log.Fatal(err)
	}
	defer emfFile.Close()
	// Get the last 300 bytes of the file
	stat, err := file.Stat()
    if err != nil {
        return nil, err
    }

    fileSize := stat.Size()
    if fileSize <= n {
        // If file is smaller than n bytes, read entire file
        return io.ReadAll(file)
    }
	return ""
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
	embeddedObjectRefs := findEmbeddedObjects(documentFile)
	referencePathsFile, err := wordDirReader.Open("word/_rels/document.xml.rels")
	if err != nil {
		log.Fatal(err)
	}
	defer referencePathsFile.Close()
	relToPathMap := getRelToPaths(referencePathsFile)
	var embeddedObjectPaths []EmbeddedObjectPaths
	for _, objRef := range embeddedObjectRefs {
		emfPath, emfOk := relToPathMap[objRef.emfID]
		pdfPath, pdfOk := relToPathMap[objRef.pdfID]
		if emfOk && pdfOk {
			embeddedObjectPaths = append(embeddedObjectPaths, EmbeddedObjectPaths{emfPath: emfPath, pdfPath: pdfPath})
			fmt.Printf("Found embedded object paths: EMF: %s, PDF: %s\n", emfPath, pdfPath)
		}
	}
}
