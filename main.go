package main

import (
	_ "database/sql"
	"fmt"
	"os"

	"github.com/xuri/excelize/v2"
	_ "modernc.org/sqlite"
)

type table [][]any

func main() {
	f := excelize.NewFile()

	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	if _, err := f.NewSheet("workers"); err != nil {
		fmt.Println(err)
		return
	}

	staff := newTable(5,2)
	staff[0][0] = "id"
	staff[0][1] = "name"
	id, name := os.Args[1], os.Args[2]
	staff[1][0] = id
	staff[1][1] = name
	insertTable(staff, f, "workers")
	if err := f.SaveAs("sheet/book.xlsx"); err != nil {
		fmt.Println(err)
		return
	}

	rows, err := f.GetRows("workers")
	if err != nil {
		fmt.Println(err)
		return
	}
	for _, row := range rows {
		for _, c := range row {
			fmt.Print(c, "\t")
		}
		fmt.Println()
	}
}

func insert(y, x int, file *excelize.File, sheet string, value any) error {
	cell, err := excelize.CoordinatesToCellName(y, x)
	if err != nil {
		return err
	}

	if err := file.SetCellValue(sheet, cell, value); err != nil {
		return err
	}

	return nil
}

func insertTable(t table, file *excelize.File, sheet string) {
	for y, row := range t {
		for x, v := range row {
			if err := insert(y+1, x+1, file, sheet, v); err != nil {
				fmt.Println(err)
				return
			}
		}
	}
}

func newTable(rows, cols int) table {
	t := make(table, rows)
	for y := range rows {
		t[y] = make([]any, cols)
	}
	return t
}
