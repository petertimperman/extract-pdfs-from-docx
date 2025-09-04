package main

import (
	"archive/zip"
	"encoding/xml"
	"flag"
	"fmt"
	"io"
	"log"
	"os"
	"regexp"
	"sort"
	"strings"
	"unicode"
	"github.com/IntelligenceX/fileconversion/ole2"
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
var (
	pdfExtPattern1 = regexp.MustCompile("\x2e\x00\x70\x00\x64\x00\x66\x00") 
	pdfExtPattern2 = regexp.MustCompile("\x2e\x00\x50\x00\x44\x00\x46\x00\x00")
)


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
					isPdf := false
					for _, attr := range element.Attr {
						if attr.Name.Local == "ProgID" && attr.Value == ADOBE_TYPE_PROG_ID {
							isPdf = true
						} else if attr.Name.Local == "id" {
							tempPdfID = attr.Value
						}
					}
					if tempPdfID != "" && isPdf {
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

func findTitleInEmf(emfFile io.Reader) string {
	// Title is stored in the last 300 bytes of the file with a specific pattern

	data , err := io.ReadAll(emfFile)
	if err != nil {
		log.Fatal(err)
	}
	content := data[len(data)-300:]
	
	// Search for the PDF extension pattern
	match1 := pdfExtPattern1.FindAllIndex(content, -1)
	match2 := pdfExtPattern2.FindAllIndex(content, -1)
	// fmt.Println(match1)
	// fmt.Println(match2)
	var loc1 []int
	var loc2 []int
	if len(match1) == 0 && len(match2) == 0 {
		// fmt.Println("No PDF extension pattern found in the last 300 bytes of the EMF file")
		return ""
	}
	if len(match1) > 0 {
		loc1 = match1[len(match1)-1]
	}else {
		loc1 = nil
	}
	if len(match2) > 0 {
		loc2 = match2[len(match2)-1]
	} else {
		loc2 = nil
	}

	if loc1 == nil && loc2 == nil {
		fmt.Println("No PDF extension pattern found in the last 300 bytes of the EMF file")
		return ""
	}
	var loc []int
	if loc1 != nil {
		loc = loc1
	} else {
		loc = loc2
	}
	// fmt.Println("Found PDF extension pattern at index:", loc[0])
	var titleCharInts []int
	startIndex := loc[0] -1 
	nullByteCounter := 0
	index := startIndex 
	for  index > 0 && nullByteCounter < 3 && (len(titleCharInts) < 30) {
		char := content[index]
		if char != 0x00 { 
			// fmt.Printf("Found title character byte: 0x%02x at index %d\n", char, index)
			titleCharInts = append([]int{int(char)}, titleCharInts...)
			nullByteCounter = 0
		} else {
			// fmt.Printf("Found null byte at index %d counter is %d \n", index, nullByteCounter)
			nullByteCounter++
		}
		index-- 
	}
	// Reverse the titleCharInts
	// fmt.Println("Extracted title character bytes (reversed):", titleCharInts)
	sort.Slice(titleCharInts, func(i, j int) bool { return i < j })
	// Convert to runes
	titleRunes := make([]rune, len(titleCharInts))
	for i, charInt := range titleCharInts {
		titleRunes[i] = rune(charInt)
	}
	// Remove non-printable characters from the title
	titleRunes = []rune(strings.TrimFunc(string(titleRunes), func(r rune) bool {
		if (!unicode.IsPrint(r) && r != '\n' && r != '\r' && r != '\t' && r != '.'){
			return true
		}
		return false
	}))
	return string(titleRunes)	
}

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
						target = "word/" + attr.Value
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

func extractPdfBytesFromBinFile() [] byte {
	
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
			// Extract the title from the emf file

			emfContentFile, err := wordDirReader.Open(emfPath)
			if err != nil {
				log.Fatal(err)
			}
			defer emfContentFile.Close()
			title := findTitleInEmf(emfContentFile)
			if title != "" {
				fmt.Printf("Extracted title: %s\n", title)
			} 
		}
	}
}
