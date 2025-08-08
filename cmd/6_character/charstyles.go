package main

import (
	"archive/zip"
	"bytes"
	"encoding/csv"
	"encoding/xml"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"
)

// FormattingType represents the type of direct formatting found
type FormattingType string

const (
	Bold        FormattingType = "Bold"
	Italic      FormattingType = "Italic"
	BoldItalic  FormattingType = "Bold Italic"
	Superscript FormattingType = "Superscript"
	Subscript   FormattingType = "Subscript"
)

// ChangeTracker tracks the changes made during conversion
type ChangeTracker struct {
	TotalChanges int
	StyleCounts  map[string]int
}

// NewChangeTracker creates a new change tracker
func NewChangeTracker() *ChangeTracker {
	return &ChangeTracker{
		TotalChanges: 0,
		StyleCounts:  make(map[string]int),
	}
}

// AddChange records a formatting change
func (ct *ChangeTracker) AddChange(styleName string) {
	ct.TotalChanges++
	ct.StyleCounts[styleName]++
}

// LibreOfficeConverter handles the conversion process
type LibreOfficeConverter struct {
	formattingMap map[FormattingType]string
	changeTracker *ChangeTracker
}

// NewLibreOfficeConverter creates a new converter instance
func NewLibreOfficeConverter() *LibreOfficeConverter {
	return &LibreOfficeConverter{
		formattingMap: make(map[FormattingType]string),
		changeTracker: NewChangeTracker(),
	}
}

// LoadStyleMappings reads the CSV file and creates the formatting map
func (loc *LibreOfficeConverter) LoadStyleMappings(filename string) error {
	file, err := os.Open(filename)
	if err != nil {
		return fmt.Errorf("failed to open file %s: %w", filename, err)
	}
	defer file.Close()

	reader := csv.NewReader(file)
	lineCount := 0

	for {
		record, err := reader.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			return fmt.Errorf("error reading CSV line %d: %w", lineCount+1, err)
		}

		lineCount++

		// Skip comment lines
		if len(record) > 0 && strings.HasPrefix(record[0], "#") {
			continue
		}

		// Ensure we have both columns
		if len(record) < 2 {
			log.Printf("Warning: skipping line %d - insufficient columns", lineCount)
			continue
		}

		formattingType := FormattingType(strings.TrimSpace(record[0]))
		characterStyle := strings.TrimSpace(record[1])

		// Validate formatting type
		if !isValidFormattingType(formattingType) {
			log.Printf("Warning: unknown formatting type '%s' on line %d", formattingType, lineCount)
			continue
		}

		loc.formattingMap[formattingType] = characterStyle
	}

	fmt.Printf("Loaded %d style mappings from %s\n", len(loc.formattingMap), filename)
	return nil
}

// isValidFormattingType checks if the formatting type is supported
func isValidFormattingType(ft FormattingType) bool {
	switch ft {
	case Bold, Italic, BoldItalic, Superscript, Subscript:
		return true
	default:
		return false
	}
}

// ODT XML structures for parsing content.xml
type ODTDocument struct {
	XMLName xml.Name           `xml:"document-content"`
	Body    ODTBody            `xml:"body"`
	Styles  ODTAutomaticStyles `xml:"automatic-styles"`
}

type ODTBody struct {
	Text ODTText `xml:"text"`
}

type ODTText struct {
	Paragraphs []ODTParagraph `xml:"p"`
}

type ODTParagraph struct {
	XMLName   xml.Name `xml:"p"`
	StyleName string   `xml:"style-name,attr,omitempty"`
	Content   []byte   `xml:",innerxml"`
}

type ODTAutomaticStyles struct {
	XMLName xml.Name   `xml:"automatic-styles"`
	Styles  []ODTStyle `xml:"style"`
}

type ODTStyle struct {
	XMLName   xml.Name           `xml:"style"`
	Name      string             `xml:"name,attr"`
	Family    string             `xml:"family,attr"`
	TextProps *ODTTextProperties `xml:"text-properties,omitempty"`
}

type ODTTextProperties struct {
	XMLName      xml.Name `xml:"text-properties"`
	FontWeight   string   `xml:"font-weight,attr,omitempty"`
	FontStyle    string   `xml:"font-style,attr,omitempty"`
	TextPosition string   `xml:"text-position,attr,omitempty"`
}

// ProcessODTFile reads an ODT file, processes it, and saves the result
func (loc *LibreOfficeConverter) ProcessODTFile(inputPath string) error {
	// Validate input file is ODT
	if !strings.HasSuffix(strings.ToLower(inputPath), ".odt") {
		return fmt.Errorf("input file must be an ODT file: %s", inputPath)
	}

	// Generate output filename
	outputPath := generateOutputFilename(inputPath)

	fmt.Printf("Processing ODT file: %s\n", inputPath)
	fmt.Printf("Output will be saved to: %s\n", outputPath)

	// Check if input file exists
	if _, err := os.Stat(inputPath); os.IsNotExist(err) {
		return fmt.Errorf("input file does not exist: %s", inputPath)
	}

	// Read ODT file (it's a ZIP archive)
	odtContent, err := loc.readODTFile(inputPath)
	if err != nil {
		return fmt.Errorf("error reading ODT file: %w", err)
	}

	// Process the content.xml
	modified, err := loc.processContentXML(odtContent["content.xml"])
	if err != nil {
		return fmt.Errorf("error processing content: %w", err)
	}

	// Update content in the ODT structure
	odtContent["content.xml"] = modified

	// Save the modified ODT file
	err = loc.saveODTFile(odtContent, outputPath)
	if err != nil {
		return fmt.Errorf("error saving ODT file: %w", err)
	}

	fmt.Printf("Successfully converted %s to %s\n", inputPath, outputPath)
	return nil
}

// generateOutputFilename creates the output filename by adding _converted before .odt
func generateOutputFilename(inputPath string) string {
	dir := filepath.Dir(inputPath)
	base := filepath.Base(inputPath)
	ext := filepath.Ext(base)
	nameWithoutExt := strings.TrimSuffix(base, ext)

	outputName := nameWithoutExt + "_converted.odt"
	return filepath.Join(dir, outputName)
}

// readODTFile reads an ODT file and returns its contents as a map
func (loc *LibreOfficeConverter) readODTFile(filePath string) (map[string][]byte, error) {
	fmt.Printf("Reading ODT file: %s\n", filePath)

	reader, err := zip.OpenReader(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open ODT file as ZIP: %w", err)
	}
	defer reader.Close()

	content := make(map[string][]byte)

	for _, file := range reader.File {
		rc, err := file.Open()
		if err != nil {
			return nil, fmt.Errorf("failed to read file %s from ODT: %w", file.Name, err)
		}

		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			return nil, fmt.Errorf("failed to read content of %s: %w", file.Name, err)
		}

		content[file.Name] = data
	}

	fmt.Printf("Successfully read ODT file with %d internal files\n", len(content))
	return content, nil
}

// processContentXML processes the content.xml and converts direct formatting to character styles
func (loc *LibreOfficeConverter) processContentXML(contentXML []byte) ([]byte, error) {
	fmt.Println("Processing content.xml for direct formatting...")

	// Parse the XML content
	var doc ODTDocument
	err := xml.Unmarshal(contentXML, &doc)
	if err != nil {
		return nil, fmt.Errorf("failed to parse content.xml: %w", err)
	}

	// Process each paragraph
	for i := range doc.Body.Text.Paragraphs {
		err := loc.processParagraph(&doc.Body.Text.Paragraphs[i], &doc.Styles)
		if err != nil {
			log.Printf("Warning: error processing paragraph %d: %v", i, err)
		}
	}

	// Marshal back to XML
	modifiedXML, err := xml.MarshalIndent(&doc, "", "  ")
	if err != nil {
		return nil, fmt.Errorf("failed to marshal modified XML: %w", err)
	}

	// Add XML declaration
	xmlHeader := []byte(`<?xml version="1.0" encoding="UTF-8"?>` + "\n")
	return append(xmlHeader, modifiedXML...), nil
}

// processParagraph processes a paragraph and its text spans for direct formatting
func (loc *LibreOfficeConverter) processParagraph(paragraph *ODTParagraph, styles *ODTAutomaticStyles) error {
	content := string(paragraph.Content)

	// Look for text spans with direct formatting
	// This is a simplified approach - real implementation would need proper XML parsing
	if strings.Contains(content, `font-weight="bold"`) ||
		strings.Contains(content, `font-style="italic"`) ||
		strings.Contains(content, `text-position`) {

		// Determine formatting type
		var formattingType FormattingType
		isBold := strings.Contains(content, `font-weight="bold"`)
		isItalic := strings.Contains(content, `font-style="italic"`)
		isSuper := strings.Contains(content, `text-position="super`)
		isSub := strings.Contains(content, `text-position="sub`)

		if isBold && isItalic {
			formattingType = BoldItalic
		} else if isBold {
			formattingType = Bold
		} else if isItalic {
			formattingType = Italic
		} else if isSuper {
			formattingType = Superscript
		} else if isSub {
			formattingType = Subscript
		} else {
			return nil // No recognized formatting
		}

		// Check if we have a mapping for this formatting type
		characterStyle, exists := loc.formattingMap[formattingType]
		if !exists {
			return nil
		}

		// Create or update character style in automatic styles
		loc.ensureCharacterStyleExists(styles, characterStyle, formattingType)

		// Replace direct formatting with character style reference
		modifiedContent := loc.replaceDirectFormattingWithStyle(content, characterStyle, formattingType)
		paragraph.Content = []byte(modifiedContent)

		// Track the change
		loc.changeTracker.AddChange(characterStyle)

		fmt.Printf("Converted direct %s formatting to style '%s'\n", formattingType, characterStyle)
	}

	return nil
}

// ensureCharacterStyleExists creates a character style if it doesn't exist
func (loc *LibreOfficeConverter) ensureCharacterStyleExists(styles *ODTAutomaticStyles, styleName string, formattingType FormattingType) {
	// Check if style already exists
	for _, style := range styles.Styles {
		if style.Name == styleName && style.Family == "text" {
			return // Style already exists
		}
	}

	// Create new character style
	newStyle := ODTStyle{
		Name:      styleName,
		Family:    "text",
		TextProps: &ODTTextProperties{},
	}

	// Set properties based on formatting type
	switch formattingType {
	case Bold:
		newStyle.TextProps.FontWeight = "bold"
	case Italic:
		newStyle.TextProps.FontStyle = "italic"
	case BoldItalic:
		newStyle.TextProps.FontWeight = "bold"
		newStyle.TextProps.FontStyle = "italic"
	case Superscript:
		newStyle.TextProps.TextPosition = "super 58%"
	case Subscript:
		newStyle.TextProps.TextPosition = "sub 58%"
	}

	// Add to automatic styles
	styles.Styles = append(styles.Styles, newStyle)
	fmt.Printf("Created character style: %s\n", styleName)
}

// replaceDirectFormattingWithStyle replaces direct formatting with character style reference
func (loc *LibreOfficeConverter) replaceDirectFormattingWithStyle(content, styleName string, formattingType FormattingType) string {
	// This is a simplified implementation
	// Real implementation would need proper XML manipulation

	switch formattingType {
	case Bold:
		// Replace <text:span text:style-name="..." fo:font-weight="bold">
		// with <text:span text:style-name="styleName">
		content = strings.ReplaceAll(content, `fo:font-weight="bold"`, "")
		content = strings.ReplaceAll(content, `font-weight="bold"`, "")
	case Italic:
		content = strings.ReplaceAll(content, `fo:font-style="italic"`, "")
		content = strings.ReplaceAll(content, `font-style="italic"`, "")
	case BoldItalic:
		content = strings.ReplaceAll(content, `fo:font-weight="bold"`, "")
		content = strings.ReplaceAll(content, `font-weight="bold"`, "")
		content = strings.ReplaceAll(content, `fo:font-style="italic"`, "")
		content = strings.ReplaceAll(content, `font-style="italic"`, "")
	case Superscript, Subscript:
		// Remove text-position attributes
		content = strings.ReplaceAll(content, `text-position="super 58%"`, "")
		content = strings.ReplaceAll(content, `text-position="sub 58%"`, "")
	}

	// Add style reference (simplified - real implementation needs proper XML handling)
	if strings.Contains(content, `text:style-name="`) {
		// Replace existing style name
		content = replaceStyleName(content, styleName)
	}

	return content
}

// replaceStyleName replaces the style name in a text span (simplified implementation)
func replaceStyleName(content, newStyleName string) string {
	// This is a very basic replacement - real implementation would use proper XML parsing
	start := strings.Index(content, `text:style-name="`)
	if start == -1 {
		return content
	}

	start += len(`text:style-name="`)
	end := strings.Index(content[start:], `"`)
	if end == -1 {
		return content
	}

	return content[:start] + newStyleName + content[start+end:]
}

// saveODTFile saves the modified ODT content to a new file
func (loc *LibreOfficeConverter) saveODTFile(content map[string][]byte, outputPath string) error {
	fmt.Printf("Saving converted ODT file to: %s\n", outputPath)

	// Create output directory if it doesn't exist
	outputDir := filepath.Dir(outputPath)
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		return fmt.Errorf("failed to create output directory: %w", err)
	}

	// Create new ZIP file
	file, err := os.Create(outputPath)
	if err != nil {
		return fmt.Errorf("failed to create output file: %w", err)
	}
	defer file.Close()

	zipWriter := zip.NewWriter(file)
	defer zipWriter.Close()

	// Write all files to the new ZIP
	for filename, data := range content {
		writer, err := zipWriter.Create(filename)
		if err != nil {
			return fmt.Errorf("failed to create file %s in ZIP: %w", filename, err)
		}

		_, err = io.Copy(writer, bytes.NewReader(data))
		if err != nil {
			return fmt.Errorf("failed to write file %s to ZIP: %w", filename, err)
		}
	}

	fmt.Printf("Successfully saved converted ODT file: %s\n", outputPath)
	return nil
}

// PrintReport displays the final conversion report
func (loc *LibreOfficeConverter) PrintReport() {
	fmt.Println("\n=== Conversion Report ===")
	fmt.Printf("Total formatting changes made: %d\n", loc.changeTracker.TotalChanges)

	if loc.changeTracker.TotalChanges == 0 {
		fmt.Println("No direct formatting found to convert.")
		return
	}

	fmt.Println("Applied character styles:")
	for styleName, count := range loc.changeTracker.StyleCounts {
		fmt.Printf("  %s: %d changes\n", styleName, count)
	}
}

func main() {
	// Check command line arguments
	if len(os.Args) < 2 {
		fmt.Printf("Usage: %s <input-document.odt> [charstyles.txt]\n", os.Args[0])
		fmt.Println("  input-document.odt: Path to the ODT document to process")
		fmt.Println("  charstyles.txt: Optional path to character styles mapping file (default: charstyles.txt)")
		os.Exit(1)
	}

	inputFile := os.Args[1]

	// Validate input file is ODT
	if !strings.HasSuffix(strings.ToLower(inputFile), ".odt") {
		log.Fatalf("Error: Input file must be an ODT file, got: %s", inputFile)
	}

	// Determine charstyles file path
	charStylesFile := "charstyles.txt"
	if len(os.Args) >= 3 {
		charStylesFile = os.Args[2]
	}

	fmt.Printf("LibreOffice ODT Character Style Converter\n")
	fmt.Printf("=========================================\n")
	fmt.Printf("Input document: %s\n", inputFile)
	fmt.Printf("Character styles file: %s\n\n", charStylesFile)

	// Create converter instance
	converter := NewLibreOfficeConverter()

	// Load style mappings from CSV file
	err := converter.LoadStyleMappings(charStylesFile)
	if err != nil {
		log.Fatalf("Error loading style mappings: %v", err)
	}

	// Process the ODT file
	err = converter.ProcessODTFile(inputFile)
	if err != nil {
		log.Fatalf("Error processing document: %v", err)
	}

	// Print final report
	converter.PrintReport()

	fmt.Println("\nConversion completed successfully!")
}
