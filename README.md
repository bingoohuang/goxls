# goxls

[![Travis CI](https://travis-ci.com/bingoohuang/goxls.svg?branch=master)](https://travis-ci.com/bingoohuang/goxls)
[![Software License](https://img.shields.io/badge/License-MIT-orange.svg?style=flat-square)](https://github.com/bingoohuang/goxls/blob/master/LICENSE.md)
[![GoDoc](https://img.shields.io/badge/godoc-reference-blue.svg?style=flat-square)](https://godoc.org/github.com/bingoohuang/goxls)
[![Coverage Status](http://codecov.io/github/bingoohuang/goxls/coverage.svg?branch=master)](http://codecov.io/github/bingoohuang/goxls?branch=master)
[![goreport](https://www.goreportcard.com/badge/github.com/bingoohuang/goxls)](https://www.goreportcard.com/report/github.com/bingoohuang/goxls)

xls package use to parse the 97-2004 microsoft xls file(".xls" suffix, NOT ".xlsx" suffix )

code forked from [extrame/xls](https://github.com/extrame/xls) and fix the test. 

```go
import (
	"fmt"
	"github.com/bingoohuang/goxls/pkg/xls"
)

func main() {
	xlFile, _ := xls.Open("example.xls", "utf-8")
	sheet := xlFile.GetSheet(0)
    
	for i := 0; i <= int(sheet.MaxRow); i++ {
		row := sheet.Row(i)
		fmt.Println("row", i + 1, "col 1", row.Col(0))
		fmt.Println("row", i + 1, "col 2", row.Col(1))
	}
}
```
