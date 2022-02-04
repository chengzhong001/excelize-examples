package main

import (
	"fmt"
	"math/rand"
	"runtime"
	"syscall"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	runtime.GC()
	startTime := time.Now()
	f := excelize.NewFile()
	streamWriter, err := f.NewStreamWriter("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}
	//设置列的的宽度
	if err := streamWriter.SetColWidth(1, 10, 15); err != nil {
		fmt.Println(err)
		return
	}

	styleID, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"#DFEBF6"}, Pattern: 1},
	})
	if err != nil {
		fmt.Println(err)
		return
	}

	// 流式设置单元格的值
	if err := streamWriter.SetRow("A1", []interface{}{
		excelize.Cell{Value: "商品订单数据报表", StyleID: styleID},
	}, excelize.RowOpts{Height: 30, Hidden: false}); err != nil {
		fmt.Println(err)
		return
	}
	header := []interface{}{}
	for _, cell := range []string{
		"订单号", "商家ID", "买家ID", "商品ID", "商品单价",
		"交易件数", "物流公司ID", "运单编号", "运单状态码", "签收状态码",
	} {
		header = append(header, cell)
	}
	header = append(header, excelize.Cell{Formula: "SUM(F3:F1000000)"})
	if err := streamWriter.SetRow("A2", header); err != nil {
		fmt.Println(err)
		return
	}

	for rowID := 3; rowID <= 1000000; rowID++ {
		row := make([]interface{}, 10)
		for colID := 0; colID < 10; colID++ {
			row[colID] = rand.Intn(640000)
		}
		cell, _ := excelize.CoordinatesToCellName(1, rowID)
		if err := streamWriter.SetRow(cell, row); err != nil {
			fmt.Println(err)
			return
		}
	}

	// 流式合并单元格
	if err := streamWriter.MergeCell("A1", "J1"); err != nil {
		fmt.Println(err)
		return
	}
	if err := streamWriter.AddTable("A2", "J1000000", `{
		"table_name": "table",
		"table_style": "TableStyleMedium2",
		"show_first_column": true,
		"show_last_column": true,
		"show_row_stripes": false,
		"show_column_stripes": true
	}`); err != nil {
		fmt.Println(err)
		return
	}

	if err := streamWriter.Flush(); err != nil {
		fmt.Println(err)
		return
	}
	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
		return
	}
	printBenchmarkInfo("generate 10 columns * 100,0000 rows:", startTime)

}

func printBenchmarkInfo(info string, startTime time.Time) {
	var memStats runtime.MemStats
	var rusage syscall.Rusage
	var bToMb = func(b uint64) uint64 {
		return b / 1024 / 1024
	}
	runtime.ReadMemStats(&memStats)
	syscall.Getrusage(syscall.RUSAGE_SELF, &rusage)
	fmt.Printf("%s\nRSS = %v\nAlloc = %v MB\nTotalAlloc = %v MB\nSys = %v MB\nNumGC = %v \nCost = %s\n",
		info, bToMb(uint64(rusage.Maxrss)), bToMb(memStats.Alloc), bToMb(memStats.TotalAlloc),
		bToMb(memStats.Sys), memStats.NumGC, time.Since(startTime))
}
