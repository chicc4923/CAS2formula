package app

import (
	"fmt"
	"log"
	"strings"

	"github.com/xuri/excelize/v2"
)

// ExcelWriter Excel对象，主要存储了Excel文件的地址
type ExcelWriter struct {
	FilePath string
}

// getActualSheetName 获取实际的工作表名称
func (ew *ExcelWriter) getActualSheetName(f *excelize.File, preferredName string) (string, error) {
	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		return "", fmt.Errorf("Excel 文件中没有工作表")
	}

	// 如果指定了工作表名称且存在，则使用它
	for _, sheet := range sheets {
		if sheet == preferredName {
			return preferredName, nil
		}
	}

	// 如果指定的工作表不存在，使用第一个工作表
	log.Printf("⚠️  警告: 工作表 '%s' 不存在，使用第一个工作表 '%s'\n", preferredName, sheets[0])
	return sheets[0], nil
}

// GetSheetList 获取所有工作表的列表
func (ew *ExcelWriter) GetSheetList() ([]string, error) {
	f, err := excelize.OpenFile(ew.FilePath)
	if err != nil {
		return nil, fmt.Errorf("打开文件失败: %v", err)
	}
	defer f.Close()

	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		return nil, fmt.Errorf("Excel 文件中没有工作表")
	}

	return sheets, nil
}

// DetectFormulaColumn 检测化学式列的位置
func (ew *ExcelWriter) DetectFormulaColumn(sheetName string) (int, string, error) {
	f, err := excelize.OpenFile(ew.FilePath)
	if err != nil {
		return 0, "", fmt.Errorf("打开文件失败: %v", err)
	}
	defer f.Close()

	actualSheetName, err := ew.getActualSheetName(f, sheetName)
	if err != nil {
		return 0, "", err
	}

	headers, err := f.GetRows(actualSheetName)
	if err != nil || len(headers) == 0 {
		return 0, "", fmt.Errorf("无法获取表头或工作表为空")
	}

	// 查找"化学式"列
	for i, header := range headers[0] {
		normalizedHeader := strings.ToLower(strings.TrimSpace(header))
		if strings.Contains(normalizedHeader, "化学式") ||
			strings.Contains(normalizedHeader, "formula") ||
			normalizedHeader == "分子式" {
			return i + 1, header, nil
		}
	}

	return 0, "", fmt.Errorf("未找到化学式列")
}

// WriteTestToFormulaColumnAuto 自动检测工作表并写入
// func (ew *ExcelWriter) WriteTestToFormulaColumnAuto(rowNumbers []int) error {
// 	f, err := excelize.OpenFile(ew.FilePath)
// 	if err != nil {
// 		return fmt.Errorf("打开文件失败: %v", err)
// 	}
// 	defer f.Close()

// 	sheets := f.GetSheetList()
// 	if len(sheets) == 0 {
// 		return fmt.Errorf("Excel 文件中没有工作表")
// 	}

// 	// 使用第一个工作表
// 	sheetName := sheets[0]
// 	log.Printf("使用工作表: %s\n", sheetName)

// 	return ew.WriteTestToFormulaColumn(sheetName, rowNumbers)
// }

// WriteTestToFormulaSafe 安全写入函数（带重试机制）
// func WriteTestToFormulaSafe(filePath string, rowNumbers []int) error {
// 	writer := &ExcelWriter{
// 		FilePath: filePath,
// 	}

// 	// 先检查文件是否存在
// 	sheets, err := writer.GetSheetList()
// 	if err != nil {
// 		return fmt.Errorf("文件检查失败: %v", err)
// 	}

// 	fmt.Printf("📋 文件中的工作表: %v\n", sheets)

// 	// 尝试在每个工作表中查找化学式列并写入
// 	for _, sheet := range sheets {
// 		colIndex, colName, err := writer.DetectFormulaColumn(sheet)
// 		if err != nil {
// 			log.Printf("⚠️  工作表 %s 中未找到化学式列: %v\n", sheet, err)
// 			continue
// 		}

// 		log.Printf("✅ 在工作表 %s 中找到化学式列: 第 %d 列 (%s)\n", sheet, colIndex, colName)

// 		// 尝试写入
// 		err = writer.WriteTestToFormulaColumn(sheet, rowNumbers)
// 		if err != nil {
// 			log.Printf("⚠️  写入工作表 %s 失败: %v\n", sheet, err)
// 			continue
// 		}

// 		log.Printf("✅ 成功向工作表 %s 写入数据\n", sheet)
// 		return nil
// 	}

// 	return fmt.Errorf("在所有工作表中都未找到化学式列或写入失败")
// }

// WriteToCell 向指定列名和行号的单元格写入数据
func WriteToCell(filePath, sheetName, columnName string, rowNumber int, value interface{}) error {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return err
	}
	defer f.Close()

	if sheetName == "" {
		sheetName = "Sheet1"
	}

	// 查找列名对应的列索引
	colIndex, err := findColumnIndex(f, sheetName, columnName)
	if err != nil {
		return err
	}

	cellName, err := excelize.CoordinatesToCellName(colIndex, rowNumber)
	if err != nil {
		return err
	}

	err = f.SetCellValue(sheetName, cellName, value)
	if err != nil {
		return err
	}

	return f.Save()
}

// 查找列名对应的列索引
func findColumnIndex(f *excelize.File, sheetName, columnName string) (int, error) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return 0, err
	}

	if len(rows) == 0 {
		return 0, fmt.Errorf("工作表为空")
	}

	headers := rows[0]
	for colIndex, header := range headers {
		if header == columnName {
			return colIndex + 1, nil
		}
	}

	return 0, fmt.Errorf("未找到列名: %s", columnName)
}
