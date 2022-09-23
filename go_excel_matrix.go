package goexcelmatrix

import "github.com/xuri/excelize/v2"

type AxisMap map[string]string

const (
	DEFAULT_SHEET = "Sheet1"
	ROW_PREFIX    = "row-"
	COL_PREFIX    = "col-"
)

type excelFile struct {
	file    *excelize.File
	axisMap AxisMap
}

type ExcelFile interface {
	SetColumns(labels []string) *excelFile
	SetRows(labels []string) *excelFile
	SetValue(row string, col string, value string) *excelFile

	Save(name ...string) error
}

func NewFile() ExcelFile {
	return &excelFile{
		file:    excelize.NewFile(),
		axisMap: AxisMap{},
	}
}

// SetColumns will pre-fill the first row of the sheet with the values provided.
// It then creates the mapping of where each value is located in the document.
//
// Example:
// 		SetColumns(["label1", "label2"], "col-") will produce {"col-label1": "B1", "col-label2": "C1"}, nil
func (ef *excelFile) SetColumns(labels []string) *excelFile {
	for column, label := range labels {
		axis, _ := excelize.CoordinatesToCellName(column+2, 1)

		if err := ef.file.SetCellValue(DEFAULT_SHEET, axis, label); err != nil {
			return ef
		}

		ef.axisMap[COL_PREFIX+label] = axis
	}
	return ef
}

// SetRows will pre-fill the first column of the sheet starting from the 2nd row.
// It then creates the mapping of where each value is located in the document.
//
// Example:
// 		SetRows(["label1", "label2"], "row-") will produce {"row-label1": "A2", "row-label2": "A3"}, nil
func (ef *excelFile) SetRows(labels []string) *excelFile {
	for row, label := range labels {
		axis, _ := excelize.CoordinatesToCellName(1, row+2)

		if err := ef.file.SetCellValue(DEFAULT_SHEET, axis, label); err != nil {
			return ef
		}

		ef.axisMap[ROW_PREFIX+label] = axis
	}
	return ef
}

// SetValue will fill the value into the given row and column label.
//
// Example:
//		SetValue("Order-1", "Status", "Failed") given that "Order1" is in A2 and "Status" in B1
//		will write "Failed" into the cell of B2
func (ef *excelFile) SetValue(row string, col string, value string) *excelFile {
	x, _, _ := excelize.CellNameToCoordinates(ef.axisMap[ROW_PREFIX+row])
	_, y, _ := excelize.CellNameToCoordinates(ef.axisMap[COL_PREFIX+col])
	axis, _ := excelize.CoordinatesToCellName(x, y)

	ef.file.SetCellValue(DEFAULT_SHEET, axis, value)

	return ef
}

// Save into an excel file. Name is optional.
func (ef *excelFile) Save(name ...string) error {
	if len(name) > 0 {
		return ef.file.SaveAs(name[0])
	}

	return ef.file.Save()
}
