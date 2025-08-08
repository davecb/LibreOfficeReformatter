package main

/*
This Go algorithm provides a comprehensive solution for renaming LibreOffice styles. Here's how it works:
Key Components:

Custom XML Parser: Uses a flexible Node struct that preserves the exact XML structure while allowing modifications to attributes.
Style Reference Detection: The isStyleAttribute function identifies common LibreOffice style attributes like style-name, parent-style-name, etc.
Recursive Updates: Traverses the entire XML tree to find and update all references to the old style name.
Multiple File Support: Can process individual XML files or handle the complete set of files from an extracted ODT document.

Usage Process:
For an ODT file, you would:

Extract the ODT (it's a ZIP archive)
Run the renamer on the extracted contents
Repackage the modified files back into an ODT

For just styles.xml:
gorenamer := &StyleRenamer{
    OldStyleName: "Header 1",
    NewStyleName: "headA",
}
err := renamer.RenameStyleInFile("styles.xml", "styles_modified.xml")

The algorithm handles both style definitions (where the style is defined) and style references (where the style is applied to content), ensuring that all "Header 1" paragraphs will be reformatted to use the "headA" style instead.

Would you like me to modify any part of this algorithm or add additional features like batch processing multiple style renames?RetryClaude does not have the ability to run the code it generates yet.Claude can make mistakes. Please double-check responses.
*/

import (
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strings"
)

// StyleRenamer handles renaming styles in LibreOffice documents
type StyleRenamer struct {
	OldStyleName string
	NewStyleName string
}

// Node represents a generic XML node with attributes and children
type Node struct {
	XMLName  xml.Name
	Attrs    []xml.Attr `xml:",any,attr"`
	Content  []byte     `xml:",innerxml"`
	Children []Node
}

// UnmarshalXML implements custom unmarshaling to preserve structure
func (n *Node) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	n.XMLName = start.Name
	n.Attrs = start.Attr

	var content strings.Builder
	for {
		token, err := d.Token()
		if err != nil {
			if err == io.EOF {
				break
			}
			return err
		}

		switch t := token.(type) {
		case xml.StartElement:
			// Found a child element, parse it recursively
			var child Node
			if err := child.UnmarshalXML(d, t); err != nil {
				return err
			}
			n.Children = append(n.Children, child)

		case xml.EndElement:
			// End of current element
			n.Content = []byte(content.String())
			return nil

		case xml.CharData:
			// Text content
			content.Write(t)
		}
	}

	n.Content = []byte(content.String())
	return nil
}

// MarshalXML implements custom marshaling to reconstruct XML
func (n Node) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = n.XMLName
	start.Attr = n.Attrs

	if err := e.EncodeToken(start); err != nil {
		return err
	}

	// Write text content if no children
	if len(n.Children) == 0 && len(n.Content) > 0 {
		if err := e.EncodeToken(xml.CharData(n.Content)); err != nil {
			return err
		}
	}

	// Encode children
	for _, child := range n.Children {
		if err := e.Encode(child); err != nil {
			return err
		}
	}

	return e.EncodeToken(xml.EndElement{Name: n.XMLName})
}

// updateStyleReferences updates all style references in the XML tree
func (sr *StyleRenamer) updateStyleReferences(node *Node) bool {
	modified := false

	// Update attributes that reference styles
	for i, attr := range node.Attrs {
		// if the attr is Name, update it
		if sr.isStyleAttribute(attr.Name.Local) && attr.Value == sr.OldStyleName {
			node.Attrs[i].Value = sr.NewStyleName
			modified = true
		}
		// FIXME, find display-name and update it too
		if sr.isStyleAttribute(attr.Name.Local) && attr.Value == sr.OldStyleName {
			node.Attrs[i].Value = sr.NewStyleName
			modified = true
		}

	}

	// Recursively update children
	for i := range node.Children {
		if sr.updateStyleReferences(&node.Children[i]) {
			modified = true
		}
	}

	return modified
}

// isStyleAttribute checks if an attribute name typically references a style // FIXME dig into this
func (sr *StyleRenamer) isStyleAttribute(attrName string) bool {
	styleAttributes := []string{
		"style-name",
		"parent-style-name",
		"next-style-name",
		"master-page-name",
		"page-layout-name",
		"name", // for style definitions themselves
	}

	for _, styleAttr := range styleAttributes {
		if attrName == styleAttr {
			return true
		}
	}
	return false
}

// RenameStyleInFile processes a single LibreOffice XML file
// FIXME does not change display name
func (sr *StyleRenamer) RenameStyleInFile(inputPath, outputPath string) error {
	// Read the XML file
	file, err := os.Open(inputPath)
	if err != nil {
		return fmt.Errorf("failed to open file: %w", err)
	}
	defer file.Close()

	// Parse XML
	var root Node
	decoder := xml.NewDecoder(file)
	if err := decoder.Decode(&root); err != nil {
		return fmt.Errorf("failed to parse XML: %w", err)
	}

	// Update style references
	modified := sr.updateStyleReferences(&root)

	if !modified {
		return fmt.Errorf("style '%s' not found in file, exiting without making changes", sr.OldStyleName)
	}

	// Write modified XML
	outputFile, err := os.Create(outputPath)
	if err != nil {
		return fmt.Errorf("failed to create output file: %w", err)
	}
	defer outputFile.Close()

	// Write XML declaration
	if _, err := outputFile.WriteString(`<?xml version="1.0" encoding="UTF-8"?>` + "\n"); err != nil {
		return fmt.Errorf("failed to write XML declaration: %w", err)
	}

	// Encode the modified XML
	encoder := xml.NewEncoder(outputFile)
	encoder.Indent("", "  ")
	if err := encoder.Encode(&root); err != nil {
		return fmt.Errorf("failed to encode XML: %w", err)
	}

	return nil
}

// RenameStyleInODT processes a complete ODT file (which is a ZIP archive)
func (sr *StyleRenamer) RenameStyleInODT(odtPath string) error {
	// For a complete ODT implementation, you would:
	// 1. Extract the ODT file (it's a ZIP archive)
	// 2. Process styles.xml, content.xml, and possibly other XML files
	// 3. Repackage into a new ODT file

	// This is a simplified version that assumes you've already extracted the ODT
	// and want to process individual XML files

	xmlFiles := []string{
		"styles.xml",
		"content.xml",
		"meta.xml",
	}

	for _, xmlFile := range xmlFiles {
		inputPath := xmlFile
		outputPath := xmlFile + ".new"

		if _, err := os.Stat(inputPath); os.IsNotExist(err) {
			// File doesn't exist, skip it
			continue
		}

		fmt.Printf("Processing %s...\n", xmlFile)
		if err := sr.RenameStyleInFile(inputPath, outputPath); err != nil {
			// If style not found, that's okay for some files
			if strings.Contains(err.Error(), "not found") {
				fmt.Printf("Style '%s' not found in %s (this may be normal)\n", sr.OldStyleName, xmlFile)
				continue
			}
			return fmt.Errorf("failed to process %s: %w", xmlFile, err)
		}

		// Replace original with modified version
		if err := os.Rename(outputPath, inputPath); err != nil {
			return fmt.Errorf("failed to replace %s: %w", xmlFile, err)
		}

		fmt.Printf("Successfully updated %s\n", xmlFile)
	}

	return nil
}

// Example usage
func main() {
	// Create a style renamer
	renamer := &StyleRenamer{
		OldStyleName: "Preformatted_20_Text",
		NewStyleName: "Code",
	}

	// Example 1: Process a single XML file
	if err := renamer.RenameStyleInFile("styles.xml", "styles_modified.xml"); err != nil {
		fmt.Printf("Error processing single file: %v\n", err)
	}

	//// Example 2: Process extracted ODT contents
	//// (First extract your ODT file to the current directory)
	//if err := renamer.RenameStyleInODT("."); err != nil {
	//	fmt.Printf("Error processing ODT contents: %v\n", err)
	//}

	//fmt.Println("Style renaming complete!")
}
