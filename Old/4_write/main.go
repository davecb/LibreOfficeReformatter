package main

import (
	"archive/zip"
	"fmt"
	"io"
	"log"
	"os"
	"regexp"
	"strings"
)

// ODT Editor struct that works with raw XML
type ODTEditor struct {
	filename   string
	contentXML []byte
	reader     *zip.ReadCloser
}

// Create a new ODT editor
func NewODTEditor(filename string) (*ODTEditor, error) {
	editor := &ODTEditor{filename: filename}

	if err := editor.load(); err != nil {
		return nil, err
	}

	return editor, nil
}

// Load the ODT document
func (e *ODTEditor) load() error {
	reader, err := zip.OpenReader(e.filename)
	if err != nil {
		return fmt.Errorf("error opening ODT file: %v", err)
	}
	e.reader = reader

	// Extract content.xml
	contentXML, err := e.extractFileFromZip("content.xml")
	if err != nil {
		return fmt.Errorf("error extracting content.xml: %v", err)
	}

	e.contentXML = contentXML
	return nil
}

// Close the editor and cleanup resources
func (e *ODTEditor) Close() {
	if e.reader != nil {
		e.reader.Close()
	}
}

// Extract a file from the ZIP archive
func (e *ODTEditor) extractFileFromZip(filename string) ([]byte, error) {
	for _, file := range e.reader.File {
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

// Change paragraph style using regex
func (e *ODTEditor) ChangeParagraphStyle(fromStyle, toStyle string) int {
	content := string(e.contentXML)

	// Pattern to match paragraph elements with specific style
	// This handles both text:p and just p elements, with various namespace prefixes
	patterns := []string{
		fmt.Sprintf(`(<[^:]*:?p\s+[^>]*style-name=")%s(")`, regexp.QuoteMeta(fromStyle)),
		fmt.Sprintf(`(<[^:]*:?p\s+[^>]*text:style-name=")%s(")`, regexp.QuoteMeta(fromStyle)),
		fmt.Sprintf(`(<[^:]*:?h\s+[^>]*style-name=")%s(")`, regexp.QuoteMeta(fromStyle)),
		fmt.Sprintf(`(<[^:]*:?h\s+[^>]*text:style-name=")%s(")`, regexp.QuoteMeta(fromStyle)),
	}

	changedCount := 0
	for _, pattern := range patterns {
		re := regexp.MustCompile(pattern)
		matches := re.FindAllString(content, -1)
		changedCount += len(matches)
		content = re.ReplaceAllString(content, "${1}"+toStyle+"${2}")
	}

	e.contentXML = []byte(content)
	return changedCount
}

// Replace text content
func (e *ODTEditor) ReplaceText(oldText, newText string) int {
	content := string(e.contentXML)
	oldContent := content

	// Simple text replacement - be careful with XML special characters
	content = strings.ReplaceAll(content, oldText, newText)

	// Count changes by comparing before and after
	changedCount := strings.Count(oldContent, oldText)

	e.contentXML = []byte(content)
	return changedCount
}

// Add a new paragraph at the end of the document
func (e *ODTEditor) AddParagraph(text, styleName string) {
	content := string(e.contentXML)

	// Find the closing tag of the text body (this might vary)
	closingTags := []string{
		"</office:text>",
		"</text>",
		"</office:body>",
		"</body>",
	}

	newParagraph := fmt.Sprintf(`<text:p text:style-name="%s">%s</text:p>`, styleName, text)

	for _, closingTag := range closingTags {
		if strings.Contains(content, closingTag) {
			content = strings.Replace(content, closingTag, newParagraph+"\n"+closingTag, 1)
			break
		}
	}

	e.contentXML = []byte(content)
}

// Get basic document statistics
func (e *ODTEditor) GetStats() map[string]int {
	content := string(e.contentXML)
	stats := make(map[string]int)

	// Count paragraphs (various possible formats)
	pPatterns := []string{`<[^:]*:?p[^>]*>`, `<text:p[^>]*>`}
	for _, pattern := range pPatterns {
		re := regexp.MustCompile(pattern)
		matches := re.FindAllString(content, -1)
		stats["paragraphs"] += len(matches)
	}

	// Count headers
	hPatterns := []string{`<[^:]*:?h[^>]*>`, `<text:h[^>]*>`}
	for _, pattern := range hPatterns {
		re := regexp.MustCompile(pattern)
		matches := re.FindAllString(content, -1)
		stats["headers"] += len(matches)
	}

	// Count tables
	tablePatterns := []string{`<[^:]*:?table[^>]*>`, `<table:table[^>]*>`}
	for _, pattern := range tablePatterns {
		re := regexp.MustCompile(pattern)
		matches := re.FindAllString(content, -1)
		stats["tables"] += len(matches)
	}

	// Rough word count (count words between > and <)
	re := regexp.MustCompile(`>[^<]+<`)
	textMatches := re.FindAllString(content, -1)
	wordCount := 0
	for _, match := range textMatches {
		text := strings.TrimSpace(match[1 : len(match)-1])
		wordCount += len(strings.Fields(text))
	}
	stats["words"] = wordCount

	return stats
}

// Display some document content (simplified)
func (e *ODTEditor) DisplayContent() {
	content := string(e.contentXML)
	fmt.Println("=== Document Content (Sample) ===")

	// Extract and display paragraph content
	re := regexp.MustCompile(`<[^:]*:?p[^>]*>([^<]*)</[^:]*:?p>`)
	matches := re.FindAllStringSubmatch(content, 10) // Limit to first 10

	for i, match := range matches {
		if len(match) > 1 && strings.TrimSpace(match[1]) != "" {
			fmt.Printf("Paragraph %d: %s\n", i+1, strings.TrimSpace(match[1]))
		}
	}

	// Extract and display header content
	re = regexp.MustCompile(`<[^:]*:?h[^>]*>([^<]*)</[^:]*:?h>`)
	matches = re.FindAllStringSubmatch(content, 5) // Limit to first 5

	for i, match := range matches {
		if len(match) > 1 && strings.TrimSpace(match[1]) != "" {
			fmt.Printf("Header %d: %s\n", i+1, strings.TrimSpace(match[1]))
		}
	}
}

// Save the modified document to a new file
func (e *ODTEditor) SaveAs(outputFilename string) error {
	// Create the new ODT file
	newFile, err := os.Create(outputFilename)
	if err != nil {
		return fmt.Errorf("error creating new file: %v", err)
	}
	defer newFile.Close()

	// Create a ZIP writer
	zipWriter := zip.NewWriter(newFile)
	defer zipWriter.Close()

	// Copy all files from original, replacing content.xml with modified version
	for _, file := range e.reader.File {
		if file.Name == "content.xml" {
			// Write the modified content.xml
			if err := e.writeFileToZip(zipWriter, "content.xml", e.contentXML); err != nil {
				return fmt.Errorf("error writing modified content.xml: %v", err)
			}
		} else {
			// Copy other files unchanged
			if err := e.copyFileToZip(zipWriter, file); err != nil {
				return fmt.Errorf("error copying file %s: %v", file.Name, err)
			}
		}
	}

	return nil
}

// Helper function to write data to a file in the ZIP archive
func (e *ODTEditor) writeFileToZip(zipWriter *zip.Writer, filename string, data []byte) error {
	writer, err := zipWriter.Create(filename)
	if err != nil {
		return err
	}

	_, err = writer.Write(data)
	return err
}

// Helper function to copy a file from source ZIP to destination ZIP
func (e *ODTEditor) copyFileToZip(zipWriter *zip.Writer, sourceFile *zip.File) error {
	// Open the source file
	sourceReader, err := sourceFile.Open()
	if err != nil {
		return err
	}
	defer sourceReader.Close()

	// Create the file in the destination ZIP
	destWriter, err := zipWriter.Create(sourceFile.Name)
	if err != nil {
		return err
	}

	// Copy the content
	_, err = io.Copy(destWriter, sourceReader)
	return err
}

// Main function demonstrating usage
func main() {
	if len(os.Args) < 2 {
		fmt.Println("Usage: go run main.go <input.odt> [output.odt]")
		fmt.Println("Example: go run main.go document.odt document_modified.odt")
		return
	}

	inputFile := os.Args[1]
	outputFile := "modified_" + inputFile
	if len(os.Args) >= 3 {
		outputFile = os.Args[2]
	}

	// Create ODT editor
	editor, err := NewODTEditor(inputFile)
	if err != nil {
		log.Fatalf("Error opening ODT file: %v", err)
	}
	defer editor.Close()

	// Display original content
	fmt.Println("=== Original Document ===")
	editor.DisplayContent()

	// Show statistics
	stats := editor.GetStats()
	fmt.Printf("\nDocument Statistics:\n")
	for key, value := range stats {
		fmt.Printf("  %s: %d\n", key, value)
	}

	// Perform some edits
	fmt.Println("\n=== Performing Edits ===")

	// Change paragraph styles (use actual style names from your document)
	styleChanges := editor.ChangeParagraphStyle("Text_20_body", "Body") // Use real style names
	fmt.Printf("Changed %d elements style\n", styleChanges)

	// Replace text
	textChanges := editor.ReplaceText("sample", "modified")
	fmt.Printf("Replaced text in %d locations\n", textChanges)

	// Add a new paragraph
	editor.AddParagraph("This paragraph was added by the Go editor.", "Body")
	fmt.Println("Added a new paragraph")

	// Save the modified document
	if err := editor.SaveAs(outputFile); err != nil {
		log.Fatalf("Error saving modified document: %v", err)
	}

	fmt.Printf("\nModified document saved as: %s\n", outputFile)
	fmt.Println("Edit completed successfully!")
}
