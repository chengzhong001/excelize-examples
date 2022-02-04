package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func main() {
	f, err := excelize.OpenFile("data.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	sheetName := "Sheet1"
	rows, _ := f.GetRows(sheetName)

	for row := 2; row <= len(rows); row++ {
		cell, err := excelize.JoinCellName("E", row)
		if err != nil {
			fmt.Println(err)
			return
		}
		if err := f.SetCellFormula(sheetName, cell, "=RANDBETWEEN(2000, 6000)"); err != nil {
			fmt.Println(err)
			return
		}
	}
	// if err := f.SetDefinedName(&excelize.DefinedName{
	// 	Name:     "源数据",
	// 	RefersTo: "Sheet1$A$1:$E$73",
	// 	Comment:  "自定义名称",
	// 	Scope:    sheetName,
	// }); err != nil {
	// 	fmt.Println(err)
	// 	return
	// }

	if err := f.AddPivotTable(&excelize.PivotTableOption{
		
	}); err!=nil{
		fmt.Println(err)
		return
	}

	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
	if err := f.Close(); err != nil {
		fmt.Println(err)
		return
	}

}
