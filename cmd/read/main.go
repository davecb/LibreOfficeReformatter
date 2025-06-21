package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"log"
	"strings"
)

// ODT XML structure definitions
type OfficeDocument struct {
	XMLName xml.Name `xml:"document-content"`
	Body    Body     `xml:"body"`
}

type Body struct {
	Text TextBody `xml:"text"`
}

type TextBody struct {
	Paragraphs []Paragraph `xml:"p"`
	Headers    []Header    `xml:"h"`
	Lists      []List      `xml:"list"`
	Tables     []Table     `xml:"table"`
}

type Paragraph struct {
	StyleName string `xml:"style-name,attr"`
	Text      string `xml:",chardata"`
	Spans     []Span `xml:"span"`
	Links     []Link `xml:"a"`
}

type Span struct {
	StyleName string `xml:"style-name,attr"`
	Text      string `xml:",chardata"`
}

type Link struct {
	Href string `xml:"href,attr"`
	Text string `xml:",chardata"`
}

type Header struct {
	Level     int    `xml:"outline-level,attr"`
	StyleName string `xml:"style-name,attr"`
	Text      string `xml:",chardata"`
	Spans     []Span `xml:"span"`
}

type List struct {
	Items []ListItem `xml:"list-item"`
}

type ListItem struct {
	Paragraphs []Paragraph `xml:"p"`
}

type Table struct {
	Name    string     `xml:"name,attr"`
	Columns []Column   `xml:"table-column"`
	Rows    []TableRow `xml:"table-row"`
}

type Column struct {
	StyleName string `xml:"style-name,attr"`
}

type TableRow struct {
	Cells []TableCell `xml:"table-cell"`
}

type TableCell struct {
	ValueType  string      `xml:"value-type,attr"`
	Value      string      `xml:"value,attr"`
	Paragraphs []Paragraph `xml:"p"`
}

// additional definitions
type Manifest struct {
	XMLName   xml.Name    `xml:"manifest"`
	FileEntry []FileEntry `xml:"file-entry"`
}

type FileEntry struct {
	FullPath  string `xml:"full-path,attr"`
	MediaType string `xml:"media-type,attr"`
}

type DocumentMeta struct {
	XMLName xml.Name `xml:"document-meta"`
	Meta    Meta     `xml:"meta"`
}

type Meta struct {
	Generator         string            `xml:"generator"`
	Title             string            `xml:"title"`
	Description       string            `xml:"description"`
	Subject           string            `xml:"subject"`
	Creator           string            `xml:"creator"`
	CreationDate      string            `xml:"creation-date"`
	Modifier          string            `xml:"modifier"`
	ModificationDate  string            `xml:"modification-date"`
	Language          string            `xml:"language"`
	EditingCycles     string            `xml:"editing-cycles"`
	EditingDuration   string            `xml:"editing-duration"`
	DocumentStatistic DocumentStatistic `xml:"document-statistic"`
}

type DocumentStatistic struct {
	TableCount     int `xml:"table-count,attr"`
	ImageCount     int `xml:"image-count,attr"`
	ObjectCount    int `xml:"object-count,attr"`
	PageCount      int `xml:"page-count,attr"`
	ParagraphCount int `xml:"paragraph-count,attr"`
	WordCount      int `xml:"word-count,attr"`
	CharacterCount int `xml:"character-count,attr"`
}

func parseODT(filename string) error {
	// Open the ODT file as a ZIP archive
	reader, err := zip.OpenReader(filename)
	if err != nil {
		return fmt.Errorf("error opening ODT file: %v", err)
	}
	defer reader.Close()

	// Find and read content.xml
	contentXML, err := extractFileFromZip(reader, "content.xml")
	if err != nil {
		return fmt.Errorf("error extracting content.xml: %v", err)
	}

	// Parse the XML content
	var doc OfficeDocument
	if err := xml.Unmarshal(contentXML, &doc); err != nil {
		return fmt.Errorf("error parsing content.xml: %v", err)
	}

	// Display the parsed content
	fmt.Println("=== ODT Document Content ===")

	// Display headers
	for i, header := range doc.Body.Text.Headers {
		headerText := header.Text //strings.TrimSpace(header.Text)
		for _, span := range header.Spans {
			headerText += span.Text //strings.TrimSpace(span.Text)
		}
		if headerText != "" {
			fmt.Printf("Header %d (Level %d): %s\n", i+1, header.Level, headerText)
		}
	}

	// Display paragraphs
	for i, para := range doc.Body.Text.Paragraphs {
		paraText := para.Text //strings.TrimSpace(para.Text)
		for _, span := range para.Spans {
			paraText += span.Text //strings.TrimSpace(span.Text)
		}
		for _, link := range para.Links {
			paraText += fmt.Sprintf(" [%s](%s)", link.Text, link.Href)
		}
		if paraText != "" {
			fmt.Printf("Paragraph %d: %s: %s\n", i+1, para.StyleName, paraText)
		}
	}

	// Display lists
	for i, list := range doc.Body.Text.Lists {
		fmt.Printf("List %d:\n", i+1)
		for j, item := range list.Items {
			for _, para := range item.Paragraphs {
				itemText := strings.TrimSpace(para.Text)
				for _, span := range para.Spans {
					itemText += strings.TrimSpace(span.Text)
				}
				if itemText != "" {
					fmt.Printf("  %d. %s\n", j+1, itemText)
				}
			}
		}
	}

	// Display tables
	for i, table := range doc.Body.Text.Tables {
		fmt.Printf("Table %d (%s):\n", i+1, table.Name)
		for rowIdx, row := range table.Rows {
			fmt.Printf("  Row %d: ", rowIdx+1)
			for _, cell := range row.Cells {
				cellText := ""
				for _, para := range cell.Paragraphs {
					paraText := strings.TrimSpace(para.Text)
					for _, span := range para.Spans {
						paraText += strings.TrimSpace(span.Text)
					}
					cellText += paraText + " "
				}
				if cell.Value != "" {
					cellText = cell.Value
				}
				fmt.Printf("[%s] ", strings.TrimSpace(cellText))
			}
			fmt.Println()
		}
	}

	return nil
}

func main() {
	if err := parseODT("2_Why_Do_We_Have_Queues.odt"); err != nil {
		log.Fatal(err)
	}
}

// Extract a specific file from the ZIP archive
func extractFileFromZip(reader *zip.ReadCloser, filename string) ([]byte, error) {
	for _, file := range reader.File {
		if file.Name == filename {
			rc, err := file.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()

			data, err := io.ReadAll(rc)
			if err != nil {
				return nil, err
			}
			return data, nil
		}
	}
	return nil, fmt.Errorf("file %s not found in archive", filename)
}

// Extract text from a paragraph including spans
func extractParagraphText(para Paragraph) string {
	text := strings.TrimSpace(para.Text)
	for _, span := range para.Spans {
		text += strings.TrimSpace(span.Text)
	}
	for _, link := range para.Links {
		text += fmt.Sprintf(" [%s](%s)", link.Text, link.Href)
	}
	return text
}

// List all files in the LibreOffice document
func listFilesInDocument(filename string) error {
	reader, err := zip.OpenReader(filename)
	if err != nil {
		return fmt.Errorf("error opening document: %v", err)
	}
	defer reader.Close()

	fmt.Printf("Files in %s:\n", filename)
	fmt.Println(strings.Repeat("-", 40))

	for _, file := range reader.File {
		fmt.Printf("  %s (size: %d bytes)\n", file.Name, file.UncompressedSize64)
	}

	return nil
}

// Get the manifest.xml content to understand document structure
func getManifest(filename string) error {
	reader, err := zip.OpenReader(filename)
	if err != nil {
		return fmt.Errorf("error opening document: %v", err)
	}
	defer reader.Close()

	manifestXML, err := extractFileFromZip(reader, "META-INF/manifest.xml")
	if err != nil {
		return fmt.Errorf("error extracting manifest.xml: %v", err)
	}

	var manifest Manifest
	if err := xml.Unmarshal(manifestXML, &manifest); err != nil {
		return fmt.Errorf("error parsing manifest.xml: %v", err)
	}

	fmt.Printf("Document manifest for %s:\n", filename)
	fmt.Println(strings.Repeat("-", 50))

	for _, entry := range manifest.FileEntry {
		fmt.Printf("  Path: %s\n", entry.FullPath)
		fmt.Printf("  Type: %s\n", entry.MediaType)
		fmt.Println()
	}

	return nil
}

// Extract metadata from meta.xml
func getDocumentMetadata(filename string) error {
	reader, err := zip.OpenReader(filename)
	if err != nil {
		return fmt.Errorf("error opening document: %v", err)
	}
	defer reader.Close()

	metaXML, err := extractFileFromZip(reader, "meta.xml")
	if err != nil {
		return fmt.Errorf("error extracting meta.xml: %v", err)
	}

	type DocumentMeta struct {
		XMLName xml.Name `xml:"document-meta"`
		Meta    Meta     `xml:"meta"`
	}

	type Meta struct {
		Generator         string            `xml:"generator"`
		Title             string            `xml:"title"`
		Description       string            `xml:"description"`
		Subject           string            `xml:"subject"`
		Creator           string            `xml:"creator"`
		CreationDate      string            `xml:"creation-date"`
		Modifier          string            `xml:"modifier"`
		ModificationDate  string            `xml:"modification-date"`
		Language          string            `xml:"language"`
		EditingCycles     string            `xml:"editing-cycles"`
		EditingDuration   string            `xml:"editing-duration"`
		DocumentStatistic DocumentStatistic `xml:"document-statistic"`
	}

	type DocumentStatistic struct {
		TableCount     int `xml:"table-count,attr"`
		ImageCount     int `xml:"image-count,attr"`
		ObjectCount    int `xml:"object-count,attr"`
		PageCount      int `xml:"page-count,attr"`
		ParagraphCount int `xml:"paragraph-count,attr"`
		WordCount      int `xml:"word-count,attr"`
		CharacterCount int `xml:"character-count,attr"`
	}

	var docMeta DocumentMeta
	if err := xml.Unmarshal(metaXML, &docMeta); err != nil {
		return fmt.Errorf("error parsing meta.xml: %v", err)
	}

	fmt.Printf("Document metadata for %s:\n", filename)
	fmt.Println(strings.Repeat("-", 50))
	fmt.Printf("Title: %s\n", docMeta.Meta.Title)
	fmt.Printf("Creator: %s\n", docMeta.Meta.Creator)
	fmt.Printf("Created: %s\n", docMeta.Meta.CreationDate)
	fmt.Printf("Modified by: %s\n", docMeta.Meta.Modifier)
	fmt.Printf("Modified: %s\n", docMeta.Meta.ModificationDate)
	fmt.Printf("Language: %s\n", docMeta.Meta.Language)
	fmt.Printf("Generator: %s\n", docMeta.Meta.Generator)
	fmt.Printf("Description: %s\n", docMeta.Meta.Description)
	fmt.Printf("Subject: %s\n", docMeta.Meta.Subject)
	fmt.Printf("Editing Cycles: %s\n", docMeta.Meta.EditingCycles)
	fmt.Printf("Editing Duration: %s\n", docMeta.Meta.EditingDuration)

	fmt.Println("\nDocument Statistics:")
	fmt.Printf("  Pages: %d\n", docMeta.Meta.DocumentStatistic.PageCount)
	fmt.Printf("  Paragraphs: %d\n", docMeta.Meta.DocumentStatistic.ParagraphCount)
	fmt.Printf("  Words: %d\n", docMeta.Meta.DocumentStatistic.WordCount)
	fmt.Printf("  Characters: %d\n", docMeta.Meta.DocumentStatistic.CharacterCount)
	fmt.Printf("  Tables: %d\n", docMeta.Meta.DocumentStatistic.TableCount)
	fmt.Printf("  Images: %d\n", docMeta.Meta.DocumentStatistic.ImageCount)
	fmt.Printf("  Objects: %d\n", docMeta.Meta.DocumentStatistic.ObjectCount)

	return nil
}

// Check if a file exists in the document archive
func fileExistsInDocument(filename, targetFile string) (bool, error) {
	reader, err := zip.OpenReader(filename)
	if err != nil {
		return false, fmt.Errorf("error opening document: %v", err)
	}
	defer reader.Close()

	for _, file := range reader.File {
		if file.Name == targetFile {
			return true, nil
		}
	}
	return false, nil
}

// Extract all text content from any LibreOffice document
func extractAllText(filename string) (string, error) {
	reader, err := zip.OpenReader(filename)
	if err != nil {
		return "", fmt.Errorf("error opening document: %v", err)
	}
	defer reader.Close()

	contentXML, err := extractFileFromZip(reader, "content.xml")
	if err != nil {
		return "", fmt.Errorf("error extracting content.xml: %v", err)
	}

	// Simple text extraction using regex or string manipulation
	// This is a basic approach - for more sophisticated parsing, use the full XML structures
	content := string(contentXML)

	// Remove XML tags (basic approach)
	var result strings.Builder
	inTag := false

	for _, char := range content {
		if char == '<' {
			inTag = true
			continue
		}
		if char == '>' {
			inTag = false
			continue
		}
		if !inTag {
			result.WriteRune(char)
		}
	}

	// Clean up whitespace
	text := result.String()
	text = strings.ReplaceAll(text, "\n", " ")
	text = strings.ReplaceAll(text, "\t", " ")

	// Remove multiple spaces
	for strings.Contains(text, "  ") {
		text = strings.ReplaceAll(text, "  ", " ")
	}

	return strings.TrimSpace(text), nil
}

// Get document type from file extension or content
func getDocumentType(filename string) string {
	if strings.HasSuffix(strings.ToLower(filename), ".odt") {
		return "Writer Document"
	}
	if strings.HasSuffix(strings.ToLower(filename), ".ods") {
		return "Calc Spreadsheet"
	}
	if strings.HasSuffix(strings.ToLower(filename), ".odp") {
		return "Impress Presentation"
	}
	if strings.HasSuffix(strings.ToLower(filename), ".odg") {
		return "Draw Graphics"
	}
	return "Unknown LibreOffice Document"
}

// Example usage of utility functions
func demonstrateUtilities() {
	filename := "document.odt"

	fmt.Printf("Document type: %s\n\n", getDocumentType(filename))

	// List all files
	if err := listFilesInDocument(filename); err != nil {
		fmt.Printf("Error listing files: %v\n", err)
	}

	fmt.Println()

	// Get manifest
	if err := getManifest(filename); err != nil {
		fmt.Printf("Error getting manifest: %v\n", err)
	}

	fmt.Println()

	// Get metadata
	if err := getDocumentMetadata(filename); err != nil {
		fmt.Printf("Error getting metadata: %v\n", err)
	}

	fmt.Println()

	// Extract all text
	if text, err := extractAllText(filename); err != nil {
		fmt.Printf("Error extracting text: %v\n", err)
	} else {
		fmt.Printf("Extracted text (first 200 chars): %s...\n",
			text[:min(200, len(text))])
	}

	// Check if specific files exist
	files := []string{"content.xml", "styles.xml", "meta.xml", "settings.xml"}
	fmt.Println("\nFile existence check:")
	for _, file := range files {
		if exists, err := fileExistsInDocument(filename, file); err != nil {
			fmt.Printf("  %s: Error checking (%v)\n", file, err)
		} else {
			fmt.Printf("  %s: %t\n", file, exists)
		}
	}
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
