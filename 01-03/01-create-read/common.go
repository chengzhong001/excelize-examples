package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

func Create() {
	// 创建一个文件，默认Sheet1
	f := excelize.NewFile()

	//创建一个工作表
	index := f.NewSheet("Sheet2")

	// 设置单元格的值，不创建Sheet2 设置值无效
	f.SetCellValue("Sheet2", "A2", "hello world")
	f.SetCellValue("Sheet1", "B2", 100)

	// 设置工作簿默认的工作表,打开直接显示的sheet页面
	f.SetActiveSheet(index)

	// 根据指定路径保存文件，重新运行会覆盖
	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}

func Read() {
	f, err := excelize.OpenFile("Book1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// 获取工作表中指定单元格的值
	cell, err := f.GetCellValue("Sheet1", "B2")
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(cell)

	rows, err := f.GetRows("Sheet1")

	if err != nil {
		fmt.Println(err)
		return
	}
	// 获取 Sheet1 上所有单元格
	for index, row := range rows {
		for _, colCell := range row {
			fmt.Println(colCell, "\t")
		}
		fmt.Println(index)
	}

}
