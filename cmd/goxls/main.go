package main

import (
	"flag"
	"fmt"
	"math"
	"os"
	"strconv"

	"github.com/bingoohuang/goxls/pkg/xls"
	"github.com/pkg/errors"
	"github.com/tealeg/xlsx/v3"
)

func main() {
	xlsFile, xlsxFile := "", ""
	flag.StringVar(&xlsFile, "file", "", "read excel file(.xls suffix)")
	flag.StringVar(&xlsxFile, "xlsx", "", "comparing excel file(.xlsx suffix)")

	flag.Parse()

	if xlsFile == "" {
		flag.PrintDefaults()
		os.Exit(0)
	}

	if xlsxFile == "" {
		printXlsFile(xlsFile)
	} else if err := CompareXlsXlsx(xlsFile, xlsxFile); err != nil {
		panic(err)
	}
}

func CompareXlsXlsx(xlsfn, xlsxfn string) error {
	xlsFile, err := xls.Open(xlsfn, "utf-8")
	if err != nil {
		return errors.Wrapf(err, "Cant open xls file: %s", xlsfn)
	}

	xlsxFile, err := xlsx.OpenFile(xlsxfn)
	if err != nil {
		return errors.Wrapf(err, "Cant open xlsx file: %s", xlsxfn)
	}

	xlsxSheetNum := len(xlsxFile.Sheets)
	xlsSheetNum := xlsFile.NumSheets()

	if xlsxSheetNum != xlsSheetNum {
		return errors.Errorf("numOfSheets, xls/xlsx MaxRow %d:%d", xlsSheetNum, xlsxSheetNum)
	}

	fmt.Printf("numOfSheets: %d\n", xlsxSheetNum)

	for sheetIndex, xlsxSheet := range xlsxFile.Sheets {
		xlsSheet := xlsFile.GetSheet(sheetIndex)
		if xlsSheet == nil {
			return errors.Errorf("Cant open xls sheet at %d", sheetIndex)
		}

		if int(xlsSheet.MaxRow+1) != xlsxSheet.MaxRow {
			return errors.Errorf("max rows diffs, xls/xlsx %d:%d", xlsSheet.MaxRow+1, xlsxSheet.MaxRow)
		}

		if xlsSheet.Name != xlsxSheet.Name {
			return errors.Errorf("sheet name diffs, xls/xlsx %s:%s", xlsSheet.Name, xlsxSheet.Name)
		}

		fmt.Printf("sheet index: %d, name: %s, total rows: %d\n", sheetIndex, xlsSheet.Name, xlsxSheet.MaxRow)

		for rowIndex := 0; rowIndex < xlsxSheet.MaxRow; rowIndex++ {
			xlsxRow, _ := xlsxSheet.Row(rowIndex)
			xlsRow := xlsSheet.Row(rowIndex)

			err2 := compareRow(xlsxSheet, xlsxRow, xlsRow, rowIndex)
			if err2 != nil {
				return err2
			}
		}
	}

	fmt.Println("there is no difference between two files")

	return nil
}

func compareRow(xlsxSheet *xlsx.Sheet, xlsxRow *xlsx.Row, xlsRow *xls.Row, rowIndex int) error {
	for cellIndex := 0; cellIndex < xlsxSheet.MaxCol; cellIndex++ {
		xlsText := xlsRow.Col(cellIndex)
		xlsxText := getXlsxCellText(xlsxRow, cellIndex)

		if xlsText == xlsxText {
			continue
		}

		if err := diffCell(xlsText, xlsxText, rowIndex, cellIndex); err != nil {
			return err
		}
	}

	return nil
}

func diffCell(xlsText string, xlsxText string, rowIndex int, cellIndex int) error {
	// try to convert to numbers
	xlsFloat, xlsErr := strconv.ParseFloat(xlsText, 64)
	xlsxFloat, xlsxErr := strconv.ParseFloat(xlsxText, 64)
	// check if numbers have no significant difference
	if xlsErr == nil && xlsxErr == nil {
		diff := math.Abs(xlsFloat - xlsxFloat)
		if diff <= 0.0000001 {
			return nil
		}

		return errors.Errorf("row/col: %d/%d, xlsx: (%s)[%d], xls: (%s)[%d], diff: %f.",
			rowIndex, cellIndex, xlsxText, len(xlsxText), xlsText, len(xlsText), diff)
	}

	return errors.Errorf("row/col: %d/%d, xlsx: (%s)[%d], xls: (%s)[%d].",
		rowIndex, cellIndex, xlsxText, len(xlsxText), xlsText, len(xlsText))
}

func getXlsxCellText(xlsxRow *xlsx.Row, cellIndex int) string {
	xlsxCell := xlsxRow.GetCell(cellIndex)
	xlsxText := xlsxCell.Value
	if xlsxText == "" && len(xlsxCell.RichText) > 0 {
		for _, r := range xlsxCell.RichText {
			xlsxText += r.Text
		}
	}
	return xlsxText
}

func printXlsFile(xlsFile string) {
	xlFile, err := xls.Open(xlsFile, "utf-8")
	if err != nil {
		panic(err)
	}

	fmt.Printf("NumSheets: %d\n", xlFile.NumSheets())

	for i := 0; i < xlFile.NumSheets(); i++ {
		sh := xlFile.GetSheet(i)
		fmt.Printf("Sheets: %d, Name: %s, TotalRows: %d\n", i, sh.Name, sh.MaxRow+1)

		printSheet(sh)
	}
}

func printSheet(sh *xls.WorkSheet) {
	for i := 0; i <= int(sh.MaxRow); i++ {
		row := sh.Row(i)
		fmt.Printf("Row %d\n", i+1)

		for j := row.FirstCol(); j <= row.LastCol(); j++ {
			fmt.Printf("Col %d: %s\n", j+1, strconv.Quote(row.Col(j)))
		}
	}
}
