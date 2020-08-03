package xls_test

import (
	"fmt"
	"testing"

	"github.com/bingoohuang/goxls/pkg/xls"
)

func TestExampleOpen(t *testing.T) {
	if xlFile, err := xls.Open("testdata/example.xls", "utf-8"); err == nil {
		fmt.Println(xlFile.Author)
	}
}

func TestExampleWorkBook_NumSheets(t *testing.T) {
	if xlFile, err := xls.Open("testdata/example.xls", "utf-8"); err == nil {
		for i := 0; i < xlFile.NumSheets(); i++ {
			sheet := xlFile.GetSheet(i)
			fmt.Println(sheet.Name)
		}
	}
}

func TestExampleOpen_GetSheet(t *testing.T) {
	xlFile, _ := xls.Open("testdata/wangpengyuan.xls", "utf-8")
	sheet := xlFile.GetSheet(0)

	for i := 0; i <= int(sheet.MaxRow); i++ {
		row := sheet.Row(i)
		fmt.Println("第", i+1, "行一列数据", row.Col(0))
		fmt.Println("第", i+1, "行二列数据", row.Col(1))
	}
}

func TestGetSheet(t *testing.T) {
	if xlFile, err := xls.Open("testdata/example.xls", "utf-8"); err == nil {
		if sheet1 := xlFile.GetSheet(0); sheet1 != nil {
			fmt.Print("Total Lines ", sheet1.MaxRow, sheet1.Name)
			col1 := sheet1.Row(0).Col(0)
			col2 := sheet1.Row(0).Col(0)
			for i := 0; i <= (int(sheet1.MaxRow)); i++ {
				row1 := sheet1.Row(i)
				col1 = row1.Col(0)
				col2 = row1.Col(1)
				fmt.Print("\n", col1, ",", col2)
			}
		}
	}
}
