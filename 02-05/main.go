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
	defer func() {
		if err := f.Close(); err != nil {
			return
		}
	}()
	f.SetDefaultFont("KaiTi")

	sheet1Name, sheet2Name, sheet3Name := "东部", "西部", "南部"
	sparklineLocation := []string{}
	sparklineRange1 := []string{}
	sparklineRange2 := []string{}
	for row := 2; row <= 10; row++ {
		location, err := excelize.JoinCellName("N", row)
		if err != nil {
			fmt.Println(err)
			return
		}
		start, err := excelize.JoinCellName("B", row)
		if err != nil {
			fmt.Println(err)
			return
		}
		end, err := excelize.JoinCellName("M", row)
		if err != nil {
			fmt.Println(err)
			return
		}
		sparklineLocation = append(sparklineLocation, location)
		sparklineRange1 = append(sparklineRange1, fmt.Sprintf("%s!%s:%s", sheet1Name, start, end))
		sparklineRange2 = append(sparklineRange2, fmt.Sprintf("%s!%s:%s", sheet2Name, start, end))
	}
	// 添加迷你图
	if err := f.AddSparkline(sheet1Name, &excelize.SparklineOption{
		Location: sparklineLocation,
		Range:    sparklineRange1,
		Markers:  false, //迷你图标记是否显示
	}); err != nil {
		fmt.Println(err)
		return
	}

	if err := f.AddSparkline(sheet2Name, &excelize.SparklineOption{
		Location: sparklineLocation,
		Range:    sparklineRange2,
		Markers:  true, //迷你图标记是否显示
		Style:    18,
		Type:     "column", //默认折线图，column柱状图
	}); err != nil {
		fmt.Println(err)
		return
	}
	//设置列的分级显示，折叠
	for col := 2; col <= 7; col++ {
		column, err := excelize.ColumnNumberToName(col)
		if err != nil {
			fmt.Println(err)
			return
		}
		if err := f.SetColOutlineLevel(sheet1Name, column, 1); err != nil {
			fmt.Println(err)
			return
		}

	}
	//设置行的分级显示，折叠
	for row := 7; row < 9; row++ {
		if err := f.SetRowOutlineLevel(sheet1Name, row, 1); err != nil {
			fmt.Println(err)
			return
		}
	}

	if err := f.SetHeaderFooter(sheet2Name, &excelize.FormatHeaderFooter{
		OddHeader:  "&C&\"Microsoft YaHei,Bold Italic\"&U&KAB7942咖啡&K000000销售数据统计表",
		OddFooter:  "&C&T",
		EvenHeader: "&C&\"Microsoft YaHei,Bold Italic\"&U&KAB7942咖啡&K000000销售数据统计表",
		EvenFooter: "&C&T",
	}); err != nil {
		fmt.Println(err)
		return
	}
	// 插入分页符
	if err := f.InsertPageBreak(sheet3Name, "G1"); err != nil {
		fmt.Println(err)
		return
	}
	// 保护工作表
	if err := f.ProtectSheet(sheet1Name, &excelize.FormatSheetProtection{
		Password:      "password",
		EditScenarios: false,
	}); err != nil {
		fmt.Println(err)
		return
	}
	// 工作表分组
	if err := f.GroupSheets([]string{sheet1Name, sheet3Name}); err != nil {
		fmt.Println(err)
		return
	}
	// 工作表取消分组
	if err := f.UngroupSheets(); err != nil {
		fmt.Println(err)
		return
	}
	if err := f.SetSheetVisible("北部", false); err != nil {
		fmt.Println(err)
		return
	}

	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
		return
	}

}
