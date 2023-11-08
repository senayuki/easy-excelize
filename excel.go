package excel

import (
	"bytes"
	"errors"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"reflect"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func ReadExcelFromUrl(urlLink string, rowStructPtr interface{}) ([]interface{}, error) {
	r, err := http.Get(urlLink)
	if err != nil {
		return nil, err
	}
	defer r.Body.Close()
	data, err := io.ReadAll(r.Body)
	if err != nil {
		return nil, err
	}

	f, err := excelize.OpenReader(bytes.NewReader(data))
	if err != nil {
		return nil, err
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	return readExcel(f, rowStructPtr)
}

func ReadExcelFromPath(path string, rowStructPtr interface{}) ([]interface{}, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}
	f, err := excelize.OpenReader(bytes.NewReader(data))
	if err != nil {
		return nil, err
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	return readExcel(f, rowStructPtr)
}
func readExcel(f *excelize.File, rowStructPtr interface{}) ([]interface{}, error) {
	headerFieldIdx := map[string]int{} // header: fieldIdx
	colIdxHeader := map[int]string{}   // colIdx: header

	// get all excel tag of rowStruct
	rowType := reflect.TypeOf(rowStructPtr)
	if rowType.Kind() != reflect.Ptr {
		return nil, errors.New("rowStructPtr must be a pointer")
	}
	elems := reflect.TypeOf(rowStructPtr).Elem()
	for i := 0; i < elems.NumField(); i++ {
		if key := elems.Field(i).Tag.Get("excel"); key == "" {
			continue
		} else {
			header := ""
			for _, tag := range strings.Split(key, ";") {
				if strings.HasPrefix(tag, "width:") {
					// NOOP
				} else {
					header = tag
				}
			}
			headerFieldIdx[header] = i
		}
	}

	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		return nil, errors.New("sheet list is empty")
	}
	sheet := sheets[0]

	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, err
	}
	result := make([]interface{}, 0, len(rows))
	for rowIdx, row := range rows {
		if rowIdx == 0 {
			// the first line will be identified as header
			for colIdx, col := range row {
				if _, ok := headerFieldIdx[col]; !ok {
					log.Println("WARNING: excel ReadExcel: unknown column header:", col)
				} else {
					colIdxHeader[colIdx] = col
				}
			}
		} else {
			newValPtr := reflect.New(rowType.Elem())
			val := newValPtr.Elem()
			for colIdx, col := range row {
				if _, ok := colIdxHeader[colIdx]; !ok {
					log.Println("WARNING: excel ReadExcel: unknown column idx:", col)
				} else {
					headerName := colIdxHeader[colIdx]
					val.Field(headerFieldIdx[headerName]).SetString(col)
				}
			}
			result = append(result, newValPtr.Interface())
		}
	}
	return result, nil
}

func ExportExcel(rowStructPtr interface{}, rows []interface{}) (*excelize.File, error) {
	var sheetName = "Sheet1"
	headers := []string{}
	valIdx := []int{}
	widthIdx := []float64{}
	// get all excel tag of rowStruct
	if reflect.TypeOf(rowStructPtr).Kind() != reflect.Ptr {
		return nil, errors.New("rowStructPtr must be a pointer")
	}
	elems := reflect.TypeOf(rowStructPtr).Elem()
	for i := 0; i < elems.NumField(); i++ {
		if key := elems.Field(i).Tag.Get("excel"); key == "" {
			continue
		} else {
			header := ""
			width := float64(0)
			for _, tag := range strings.Split(key, ";") {
				if strings.HasPrefix(tag, "width:") {
					width, _ = strconv.ParseFloat(strings.TrimPrefix(tag, "width:"), 64)
				} else {
					header = tag
				}
			}
			headers = append(headers, header)
			widthIdx = append(widthIdx, width)
			valIdx = append(valIdx, i)
		}
	}

	f := excelize.NewFile()
	sheetIndex, err := f.GetSheetIndex(sheetName)
	if err != nil {
		return nil, err
	}
	for idx, w := range widthIdx {
		if w != 0 {
			f.SetColWidth(sheetName, getColumnName(idx), getColumnName(idx), w)
		}
	}
	rowCounter := 1
	// header
	for i, header := range headers {
		col := getColumnName(i)
		cellId := fmt.Sprintf("%s%d", col, rowCounter) // cellId like A1 A2
		f.SetCellValue(sheetName, cellId, header)
	}
	// data
	for _, row := range rows {
		rowCounter += 1
		if reflect.TypeOf(row).Kind() != reflect.Ptr {
			return nil, errors.New("rows input must be a pointer")
		}
		vals := reflect.ValueOf(row).Elem()
		for colIdx, reflectIdx := range valIdx {
			colName := getColumnName(colIdx)
			cellName := fmt.Sprintf("%s%d", colName, rowCounter) // cellName like A1 A2
			f.SetCellValue(sheetName, cellName, vals.Field(reflectIdx).Interface())
		}
	}
	f.SetActiveSheet(sheetIndex)
	return f, nil
}

// get column like A B C
// i start at 0 instead of 1, which means it returns A when i=0
func getColumnName(i int) string {
	col, _ := excelize.ColumnNumberToName(i + 1)
	return col
}
