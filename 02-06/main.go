package main

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	f, err := excelize.OpenFile("Book1.xlsx", excelize.Options{
		Password: "password",
	})
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// 获取sheet的索引和名称
	for sheedIdx, sheetName := range f.GetSheetList() {
		fmt.Printf("sheet index: %d, sheet name: %s\n", sheedIdx, sheetName)
	}
	activeSheetName := f.GetSheetList()[f.GetActiveSheetIndex()]
	fmt.Println("active sheet name is: ", activeSheetName)

	// 读取合并单元格
	mergedCells, err := f.GetMergeCells(activeSheetName)
	if err != nil {
		fmt.Println(err)
		return
	}
	// 遍历合并单元格
	for _, mergedCell := range mergedCells {
		fmt.Printf("merged cell range: %s:%s, value: %s\n", mergedCell.GetStartAxis(),
			mergedCell.GetEndAxis(), mergedCell.GetCellValue())
	}
	//搜索指定值的单元格
	searchResult, err := f.SearchSheet(activeSheetName, "75")
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Printf("search of 75: %s\n", strings.Join(searchResult, ","))
	// 开启正则表达式搜索
	searchResult, err = f.SearchSheet(activeSheetName, "^7", true)
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Printf("search of ^7: %s\n", strings.Join(searchResult, ","))

	rows, err := f.GetRows(activeSheetName)
	if err!=nil{
		fmt.Println(err)
		return
	}
	for _, row := range rows{
		for _, cell := range row{
			fmt.Printf("%s\t", cell)
		}
		fmt.Println()
	}
	for sheetName, comments := range f.GetComments() {
		for _, comment := range comments {
			fmt.Printf("sheet name: %s, author: %s, text: %s\n", sheetName, comment.Author, comment.Text)
		}
	}
	width, err := f.GetColWidth(activeSheetName, "A")
	if err!=nil{
		fmt.Println(err)
		return
	}
	fmt.Println("width of column A:", width)
	width, err = f.GetColWidth(activeSheetName, "G")
	if err!=nil{
		fmt.Println(err)
		return
	}
	fmt.Println("width of column G:", width)

	height, err := f.GetRowHeight(activeSheetName, 1)
	if err!= nil{
		fmt.Println(err)
		return
	}
	fmt.Println("height of row 1:", height)

	pictureName, content, err := f.GetPicture(activeSheetName, "G8")
	if err!= nil{
		fmt.Println(err)
		return
	}
	fmt.Printf("picture's name: %s, size: %dbytes\n", pictureName, len(content))
	fmt.Println("workbook default font is:", f.GetDefaultFont())

}
