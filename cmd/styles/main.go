package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"os"
	"sort"
	"strings"
)

// Style represents a LibreOffice style element
type Style struct {
	XMLName     xml.Name `xml:"style"`
	Name        string   `xml:"name,attr"`
	DisplayName string   `xml:"display-name,attr"`
	Family      string   `xml:"family,attr"`
}

// Styles represents the root styles element
type Styles struct {
	XMLName xml.Name `xml:"document-styles"`
	Styles  []Style  `xml:"styles>style"`
}

// AutomaticStyles represents automatic styles in content.xml
type AutomaticStyles struct {
	XMLName xml.Name `xml:"document-content"`
	Styles  []Style  `xml:"automatic-styles>style"`
}

// StyleInfo holds the display name and family for a style
type StyleInfo struct {
	DisplayName string
	Family      string
}

func extractStyleMapping(odfFile string) (map[string]StyleInfo, error) {
	// Open the ODF file as a ZIP archive
	reader, err := zip.OpenReader(odfFile)
	if err != nil {
		return nil, fmt.Errorf("failed to open file '%s': %v", odfFile, err)
	}
	defer reader.Close()

	mapping := make(map[string]StyleInfo)

	// Extract styles from styles.xml
	if err := extractFromFile(&reader.Reader, "styles.xml", mapping); err != nil {
		fmt.Printf("Warning: Could not extract from styles.xml: %v\n", err)
	}

	// Extract automatic styles from content.xml
	if err := extractFromFile(&reader.Reader, "content.xml", mapping); err != nil {
		fmt.Printf("Warning: Could not extract from content.xml: %v\n", err)
	}

	if len(mapping) == 0 {
		return nil, fmt.Errorf("no styles found in the document")
	}

	return mapping, nil
}

func extractFromFile(reader *zip.Reader, filename string, mapping map[string]StyleInfo) error {
	var file *zip.File
	for _, f := range reader.File {
		if f.Name == filename {
			file = f
			break
		}
	}

	if file == nil {
		return fmt.Errorf("'%s' not found in archive", filename)
	}

	rc, err := file.Open()
	if err != nil {
		return fmt.Errorf("failed to open '%s': %v", filename, err)
	}
	defer rc.Close()

	decoder := xml.NewDecoder(rc)

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch se := token.(type) {
		case xml.StartElement:
			if se.Name.Local == "style" && se.Name.Space != "" {
				var style Style
				if err := decoder.DecodeElement(&style, &se); err == nil {
					if style.Name != "" {
						displayName := style.DisplayName
						if displayName == "" {
							displayName = style.Name
						}
						mapping[style.Name] = StyleInfo{
							DisplayName: displayName,
							Family:      style.Family,
						}
					}
				}
			}
		}
	}

	return nil
}

func printStyleMappings(mapping map[string]StyleInfo) {
	// Group styles by family
	families := make(map[string][]struct {
		InternalName string
		DisplayName  string
	})

	for internalName, styleInfo := range mapping {
		family := styleInfo.Family
		if family == "" {
			family = "unknown"
		}
		families[family] = append(families[family], struct {
			InternalName string
			DisplayName  string
		}{
			InternalName: internalName,
			DisplayName:  styleInfo.DisplayName,
		})
	}

	// Sort families for consistent output
	var familyNames []string
	for family := range families {
		familyNames = append(familyNames, family)
	}
	sort.Strings(familyNames)

	// Display the mappings organized by style family
	for _, family := range familyNames {
		styles := families[family]
		fmt.Printf("\n%s STYLES:\n", strings.ToUpper(family))
		fmt.Println(strings.Repeat("-", 30))

		// Sort styles within each family
		sort.Slice(styles, func(i, j int) bool {
			return styles[i].InternalName < styles[j].InternalName
		})

		for _, style := range styles {
			fmt.Printf("  %-15s -> %s\n", style.InternalName, style.DisplayName)
		}
	}

	fmt.Printf("\nTotal styles found: %d\n", len(mapping))
}

func main() {
	if len(os.Args) != 2 {
		fmt.Println("Usage: go run extract_styles.go <path_to_odf_file>")
		fmt.Println("Example: go run extract_styles.go document.odt")
		os.Exit(1)
	}

	odfFile := os.Args[1]

	fmt.Printf("Extracting style mappings from: %s\n", odfFile)
	fmt.Println(strings.Repeat("-", 50))

	mapping, err := extractStyleMapping(odfFile)
	if err != nil {
		fmt.Printf("Error: %v\n", err)
		os.Exit(1)
	}

	printStyleMappings(mapping)
}

