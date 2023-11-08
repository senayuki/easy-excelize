package excel_test

import (
	"errors"
	"reflect"
	"testing"

	excel "github.com/senayuki/easy-excelize"
)

var testData = []interface{}{
	&ExportTest{
		ID:     "1",
		Name:   "赵",
		Mobile: "111111111",
		Age:    "10",
	},
	&ExportTest{
		ID:     "2",
		Name:   "钱",
		Mobile: "222222222",
		Age:    "20",
	},
	&ExportTest{
		ID:     "3",
		Name:   "孙",
		Mobile: "333333333",
		Age:    "30",
	},
	&ExportTest{
		ID:     "4",
		Name:   "李",
		Mobile: "444444444",
		Age:    "40",
	},
}

type ExportTest struct {
	ID     string `excel:"ID"`
	Name   string `excel:"姓名"`
	Mobile string `excel:"手机号;width:16"`
	Age    string `excel:"年龄"`
}

func TestExportExcel(t *testing.T) {
	f, err := excel.ExportExcel(&ExportTest{}, testData)
	if err != nil {
		t.Error(err)
	}
	if err := f.SaveAs("Book1.xlsx"); err != nil {
		t.Error(err)
	}
}

func TestReadExcel(t *testing.T) {
	result, err := excel.ReadExcelFromPath("Book1.xlsx", &ExportTest{})
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(result, testData) {
		t.Fatal(errors.New("result != testData"))
	}
}
