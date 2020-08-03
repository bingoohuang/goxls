package xls

import (
	"fmt"
	"io/ioutil"
	"math"
	"os"
	"path"
	"path/filepath"
	"strconv"
	"strings"
	"testing"

	"github.com/tealeg/xlsx/v3"
)

func TestIssue47(t *testing.T) {
	files, err := ioutil.ReadDir("testdata")
	if err != nil {
		t.Fatalf("Cant read testdata directory contents: %s", err)
	}
	for _, f := range files {
		if filepath.Ext(f.Name()) == ".xlsx" {
			xlsxfn := f.Name()
			xlsfn := strings.TrimSuffix(xlsxfn, ".xlsx") + ".xls"

			xlsFullFn := path.Join("testdata", xlsfn)
			xlsxFullFn := path.Join("testdata", xlsxfn)

			if _, err := os.Stat(xlsFullFn); err != nil {
				continue
			}

			fmt.Printf("Compare %s with %s\n", xlsfn, xlsxfn)
			err := CompareXlsXlsx(xlsFullFn, xlsxFullFn)
			if err != "" {
				t.Fatalf("XLS file %s an XLSX file are not equal: %s", xlsfn, err)
			}

		}
	}
}

// CompareXlsXlsx compares xls and xlsx files
func CompareXlsXlsx(xlsfilepathname string, xlsxfilepathname string) string {
	xlsFile, err := Open(xlsfilepathname, "utf-8")
	if err != nil {
		return fmt.Sprintf("Cant open xls file: %s", err)
	}

	xlsxFile, err := xlsx.OpenFile(xlsxfilepathname)
	if err != nil {
		return fmt.Sprintf("Cant open xlsx file: %s", err)
	}

	fmt.Printf("numOfSheets, xlsx/xlsx MaxRow %d:%d\n", len(xlsxFile.Sheets), xlsFile.NumSheets())

	for sheetIndex, xlsxSheet := range xlsxFile.Sheets {
		xlsSheet := xlsFile.GetSheet(sheetIndex)
		if xlsSheet == nil {
			return "Cant get xls sheet"
		}

		fmt.Printf("sheetIndex %d %s:%s, xlsx/xlsx MaxRow %d:%d\n",
			sheetIndex, xlsSheet.Name, xlsSheet.Name,
			xlsxSheet.MaxRow, xlsSheet.MaxRow+1)

		for rowIndex := 0; rowIndex < xlsxSheet.MaxRow; rowIndex++ {
			xlsxRow, _ := xlsxSheet.Row(rowIndex)
			xlsRow := xlsSheet.Row(rowIndex)

			for cellIndex := 0; cellIndex < xlsxSheet.MaxCol; cellIndex++ {
				xlsxCell := xlsxRow.GetCell(cellIndex)
				xlsxText := xlsxCell.Value
				if xlsxText == "" && len(xlsxCell.RichText) > 0 {
					for _, r := range xlsxCell.RichText {
						xlsxText += r.Text
					}
				}

				xlsText := xlsRow.Col(cellIndex)

				if xlsText == xlsxText {
					continue
				}

				// try to convert to numbers
				xlsFloat, xlsErr := strconv.ParseFloat(xlsText, 64)
				xlsxFloat, xlsxErr := strconv.ParseFloat(xlsxText, 64)
				// check if numbers have no significant difference
				if xlsErr == nil && xlsxErr == nil {
					diff := math.Abs(xlsFloat - xlsxFloat)
					if diff <= 0.0000001 {
						continue
					}

					return fmt.Sprintf("sheetIndex:%d, row/col: %d/%d, xlsx: (%s)[%d], xls: (%s)[%d], numbers difference: %f.",
						sheetIndex, rowIndex, cellIndex, xlsxText, len(xlsxText), xlsText, len(xlsText), diff)
				}

				return fmt.Sprintf("sheetIndex:%d, row/col: %d/%d, xlsx: (%s)[%d], xls: (%s)[%d].",
					sheetIndex, rowIndex, cellIndex, xlsxText, len(xlsxText), xlsText, len(xlsText))
			}
		}
	}

	return ""
}
