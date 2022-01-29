package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func main() {
	f := excelize.NewFile()
	sheetName := "成绩单"
	f.SetSheetName("Sheet1", sheetName)
	data := [][]interface{}{
		{"考试成绩统计表"},
		{"考试名称:期中考试", nil, nil, nil, "基础科目", nil, nil, "理科科目"},
		{"序号", "学号", "姓名", "班级", "数学", "英语", "语文", "化学", "生物", "物理", "总分"},
		{"1", "10001", "学生A", "1班", 93, 80, 89, 86, 57, 77},
		{"2", "10002", "学生B", "1班", 65, 72, 91, 75, 66, 90},
		{"3", "10003", "学生C", "2班", 92, 99, 89, 90, 79, 69},
		{"4", "10004", "学生D", "1班", 72, 69, 71, 82, 75, 83},
		{"5", "10005", "学生E", "2班", 81, 93, 59, 76, 66, 90},
		{"6", "10006", "学生F", "2班", 92, 90, 87, 88, 92, 70},
	}
	for i, row := range data {
		startCell, err := excelize.JoinCellName("A", i+1)
		if err != nil {
			fmt.Println(err)
			return
		}
		// 按行设置单元格的数据
		if err := f.SetSheetRow(sheetName, startCell, &row); err != nil {
			fmt.Println(err)
			return
		}
	}
	// 计算单元格总数
	formulaType, ref := excelize.STCellFormulaTypeShared, "K4:K9"
	if err := f.SetCellFormula(sheetName, "K4", "=SUM(E4:J4)",
		excelize.FormulaOpts{Ref: &ref, Type: &formulaType}); err != nil {
		fmt.Println(err)
		return
	}

	mergeCellRanges := [][]string{{"A1", "K1"}, {"A2", "D2"}, {"E2", "G2"}, {"H2", "J2"}}
	for _, ranges := range mergeCellRanges {
		if err := f.MergeCell(sheetName, ranges[0], ranges[1]); err != nil {
			fmt.Println(err)
		}
	}
	// 定义单元格的样式
	style1, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#DFEBF6"}, Pattern: 1},
	})

	if err != nil {
		fmt.Println(err)
		return
	}

	// 应用定义的单元格样式
	if f.SetCellStyle(sheetName, "A1", "A1", style1); err != nil {
		fmt.Println(err)
		return
	}
	// 定义单元格的样式
	style2, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center"},
	})
	if err != nil {
		fmt.Println(err)
		return
	}

	// 应用定义的单元格样式
	for _, cell := range []string{"A2", "E2", "H2"} {
		if f.SetCellStyle(sheetName, cell, cell, style2); err != nil {
			fmt.Println(err)
			return
		}
	}
	// 设置列的宽度
	if err := f.SetColWidth(sheetName, "D", "K", 7); err != nil {
		fmt.Println(err)
		return
	}

	// 设置筛选框表格
	if err := f.AddTable(sheetName, "A3", "K9", `{
		"table_name": "table",
		"table_style": "TableStyleLight2"
	}`); err != nil {
		fmt.Println(err)
		return
	}
	// 添加图表工作表
	if err := f.AddChart(sheetName, "A10", `{
		"type": "col",
		"series": [
			{
				"name": "成绩单!$A$2",
				"categories": "成绩单!$C$4:$C$9",
				"values": "成绩单!$J$4:$J$9"
			}
		],
		"format": {
			"x_scale": 1.3,
			"x_offset": 10,
			"y_offset": 20
		},
		"legend": {
			"none": true 
		},
		"title": {
			"name": "成绩单"
		}
	}`); err != nil {
		fmt.Println(err)
	}

	// 添加图表
	if err := f.AddChartSheet("统计图", `{
		"type": "col",
		"series": [
			{
				"name": "成绩单!$A$2",
				"categories": "成绩单!$C$4:$C$9",
				"values": "成绩单!$J$4:$J$9"
			}
		],
		"legend": {
			"none": true 
		},
		"title": {
			"name": "成绩单"
		},
		"plotarea": {
			"show_val": true
		}
	}`); err != nil {
		fmt.Println(err)
		return
	}

	// 设置背景
	if err := f.SetSheetBackground(sheetName, "watermark.jpg"); err != nil {
		fmt.Println(err)
	}

	// 设置条件格式
	red, err := f.NewConditionalStyle(`{
		"font": {
			"color": "#9A0511"
		},
		"fill": {
			"type" :"pattern",
			"color": ["#FEC7CE"],
			"pattern": 1
		}
	}`)
	if err != nil {
		fmt.Println(err)
		return
	}
	bottomCond := fmt.Sprintf(`[
		{
			"type": "bottom",
			"criteria": "=",
			"value": "1",
			"format": %d
		}
	]`, red)


	green, err := f.NewConditionalStyle(`{
		"font": {
			"color": "#09600B"
		},
		"fill": {
			"type" :"pattern",
			"color": ["#C7EECF"],
			"pattern": 1
		}
	}`)
	if err != nil {
		fmt.Println(err)
		return
	}

	topCond := fmt.Sprintf(`[
		{
			"type": "top",
			"criteria": "=",
			"value": "1",
			"format": %d
		}
	]`, green)

	for _, col := range []string{"E", "F", "G", "H", "I", "J"} {
		ref := fmt.Sprintf("%s4:%s9", col, col)
		if err := f.SetConditionalFormat(sheetName, ref, bottomCond); err != nil {
			fmt.Println(err)
		}
		if err := f.SetConditionalFormat(sheetName, ref, topCond); err != nil {
			fmt.Println(err)
		}
	}

	
	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}

}
