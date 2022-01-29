package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"os"
	"strconv"

	"github.com/xuri/excelize/v2"
)

func main() {
	csvFile, err := os.Open("MSFT.csv")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer csvFile.Close()
	reader := csv.NewReader(csvFile)

	f := excelize.NewFile()
	sheetName := "Sheet1"
	row := 1
	for {
		record, err := reader.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			fmt.Println(err)
			break
		}
		cell, err := excelize.CoordinatesToCellName(1, row)
		if err != nil {
			fmt.Println(err)
			break
		}
		if row == 1 {
			if err := f.SetSheetRow(sheetName, cell, &record); err != nil {
				fmt.Println(err)
			}
			row++
			continue
		}
		numbers, err := convertSlice(record)
		if err != nil {
			fmt.Println(err)
			break
		}
		if err := f.SetSheetRow(sheetName, cell, &numbers); err != nil {
			fmt.Println(err)
			break
		}
		row++
	}
	//修改单元格数据
	if err := f.SetCellValue(sheetName, "A1", "日期年月日"); err != nil {
		fmt.Println(err)
		return
	}
	// 设置单元格为字符串
	if err := f.SetCellStr(sheetName, "B1", "收盘价"); err != nil {
		fmt.Println(err)
		return
	}
	// 设置整行单元格
	if err := f.SetSheetRow(sheetName, "C1", &[]interface{}{"开盘价", "最高价", "最低价", "收盘调价", "交易量", "涨跌幅百分比"}); err != nil {
		fmt.Println(err)
		return
	}
	// 保留两位小数
	style1, err := f.NewStyle(&excelize.Style{NumFmt: 2})
	if err != nil {
		fmt.Println(err)
		return
	}
	if err := f.SetColStyle(sheetName, "B:F", style1); err != nil {
		fmt.Println(err)
	}
	// 长数字设置千位逗号分隔
	style2, err := f.NewStyle(&excelize.Style{NumFmt: 3})
	if err != nil {
		fmt.Println(err)
		return
	}
	if err := f.SetColStyle(sheetName, "G", style2); err != nil {
		fmt.Println(err)
		return
	}
	// 设置列宽
	if err := f.SetColWidth(sheetName, "A", "G", 11); err != nil {
		fmt.Println(err)
		return
	}
	// 隐藏列
	if err := f.SetColVisible(sheetName, "F", false); err != nil {
		fmt.Println(err)
		return
	}
	//插入行
	if err := f.InsertRow(sheetName, 1); err != nil {
		fmt.Println(err)
		return
	}
	//合并单元格
	if err := f.MergeCell(sheetName, "A1", "H1"); err != nil {
		fmt.Println(err)
		return
	}
	//设置单元格富文本格式
	if err := f.SetCellRichText(sheetName, "A1", []excelize.RichTextRun{
		{
			Text: "MSFT\r\n",
			Font: &excelize.Font{
				Bold:   true,
				Color:  "2354e8",
				Size:   20,
				Family: "Times New Roman",
			},
		}, {
			Text: "近五年数据",
			Font: &excelize.Font{
				Family: "Microsoft YaHei",
			},
		},
	}); err != nil {
		fmt.Println(err)
		return
	}
	// 设置单元格换行样式
	style3, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			WrapText:   true, //自动换行
			Horizontal: "center",
			Vertical:   "center",
		},
	})
	if err != nil {
		fmt.Println(err)
		return
	}
	// 应用单元格的换行样式
	if err := f.SetCellStyle(sheetName, "A1", "A1", style3); err != nil {
		fmt.Println(err)
		return
	}
	//设置行高
	if err := f.SetRowHeight(sheetName, 1, 60); err != nil {
		fmt.Println(err)
		return
	}

	if err := f.SetCellValue(sheetName, "I1", "数据来源: finance.yahoo.com"); err != nil {
		fmt.Println(err)
		return
	}
	//设置单元格超链接
	if err := f.SetCellHyperLink(sheetName, "I1", "https://finance.yahoo.com", "External"); err != nil {
		fmt.Println(err)
		return
	}
	// 设置超链接样式
	style4, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Color:     "#1265BE",
			Underline: "single",
		},
	})
	if err != nil {
		fmt.Println(err)
		return
	}
	if err := f.SetCellStyle(sheetName, "I1", "I1", style4); err != nil {
		fmt.Println(err)
	}
	// // 创建一个新的sheet页
	// sheetIdx := f.NewSheet("走势图")
	// f.SetActiveSheet(sheetIdx)

	// if err := f.AddChart("走势图", "A1", `{
	// 	"type": "line",
	// 	"series": [{
	// 		"name": "Sheet1!$E$2",
	// 		"categories": "Sheet1!$A$3:$A1261",
	// 		"values": "Sheet1!$E$3:$E$1261"
	// 	}],
	// 	"format":{
	// 		"x_scale": 1.6,
	// 		"y_scale": 1.5,
	// 		"x_offset: 15,
	// 		"y_offset: 10
	// 	},
	// 	"title": {
	// 		"name": "收盘价"
	// 	},
	// 	"legend": {
	// 		"none": true
	// 	}
	// }`); err != nil {
	// 	fmt.Println(err)
	// 	return
	// }

	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
		return
	}
}

func convertSlice(record []string) (numbers []interface{}, err error) {
	for _, arg := range record {
		var n float64
		if n, err = strconv.ParseFloat(arg, 64); err == nil {
			numbers = append(numbers, n)
			continue
		}
		numbers = append(numbers, arg)
	}
	return
}
