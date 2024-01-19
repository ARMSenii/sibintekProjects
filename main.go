package main

import (
	"encoding/json"
	"fmt"
	"log"
	"os"
	"strconv"

	"github.com/tealeg/xlsx"
)

type XlsxData struct {
	Worksheets []Worksheet `json:"worksheets"`
}

type Worksheet struct {
	WorksheetName string `json:"worksheetName"`
	Rows          []Row  `json:"rows"`
}

type Row struct {
	RowIndex string `json:"rowIndex"`
	Cells    []Cell `json:"cells"`
}

type Cell struct {
	ColumnIndex string      `json:"columnIndex"`
	DataType    string      `json:"dataType"`
	CellFormula string      `json:"cellFormula"`
	Style       string      `json:"style"`
	Value       interface{} `json:"value"`
	//MergedCells []int       `json:"mergedCells,omitempty"`
	MergedCells int `json:"mergedCells"`
}

func main() {
	if len(os.Args) < 2 {
		fmt.Println("путь к excel")
		return
	}

	filePath := os.Args[1]

	xlFile, err := xlsx.OpenFile(filePath)
	if err != nil {
		log.Fatal(err)
	}

	var xlsxData XlsxData

	// Проходим по всем листам в файле Excel
	for _, sheet := range xlFile.Sheets {
		var worksheet Worksheet
		worksheet.WorksheetName = sheet.Name

		// Проходим по всем строкам на листе
		for _, row := range sheet.Rows {
			var rowData Row

			// Проходим по всем ячейкам в строке
			for a, cell := range row.Cells {
				columnIndex := xlsx.ColIndexToLetters(a)
				cellData := Cell{
					ColumnIndex: columnIndex,
					DataType:    strconv.Itoa(int(cell.Type())),
					CellFormula: cell.Formula(),
					Style:       cell.GetStyle().Fill.PatternType,
					Value:       cell.Value,
					MergedCells: cell.HMerge,
				}

				rowData.Cells = append(rowData.Cells, cellData)
			}

			worksheet.Rows = append(worksheet.Rows, rowData)
		}

		xlsxData.Worksheets = append(xlsxData.Worksheets, worksheet)
	}

	// Преобразуем структуру xlsxData в JSON и выводим в консоль
	jsonData, err := json.MarshalIndent(xlsxData, "", "    ")
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println(string(jsonData))
}
