package main

import (
	"fmt"
	"github.com/xiaokeng7788/DataMergeExcel/common"
	"testing"
)

func TestGetExcelAppointIndexData(t *testing.T) {
	res, err := common.GetExcelAppointIndexData("1.xlsx", "Sheet1", "通话id", 1)
	if err != nil {
		t.Error(err)
	}
	fmt.Println(res)
}
