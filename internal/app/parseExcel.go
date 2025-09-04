package app

import (
	"fmt"
	"log"
	"os"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// ExcelProcessor excel对象
type ExcelProcessor struct {
	FilePath string
}

// ProcessEmptyChemicalFormulas 处理化学式为空的记录
func (ep *ExcelProcessor) ProcessEmptyChemicalFormulas() ([]int, int, error) {
	startTime := time.Now()
	log.Printf("开始处理文件: %s\n", ep.FilePath)

	f, err := excelize.OpenFile(ep.FilePath)
	if err != nil {
		return nil, 0, fmt.Errorf("打开文件失败: %v", err)
	}
	defer f.Close()

	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		return nil, 0, fmt.Errorf("Excel 文件中没有工作表")
	}

	var allEmptyRows []int
	totalEmptyCount := 0

	// 处理所有工作表
	for _, sheetName := range sheets {
		log.Printf("\n处理工作表: %s\n", sheetName)

		emptyRows, emptyCount, err := ep.processSheet(f, sheetName)
		if err != nil {
			log.Printf("警告: 处理工作表 %s 时出错: %v", sheetName, err)
			continue
		}

		allEmptyRows = append(allEmptyRows, emptyRows...)
		totalEmptyCount += emptyCount
	}

	elapsed := time.Since(startTime)
	log.Printf("\n处理完成! 耗时: %v\n", elapsed)

	return allEmptyRows, totalEmptyCount, nil
}

// processSheet 处理单个工作表
func (ep *ExcelProcessor) processSheet(f *excelize.File, sheetName string) ([]int, int, error) {
	// 使用 GetRows 获取所有数据
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, 0, fmt.Errorf("读取行数据失败: %v", err)
	}

	if len(rows) == 0 {
		log.Printf("  工作表 %s 为空\n", sheetName)
		return nil, 0, nil
	}

	log.Printf("  总行数: %d (包含表头)\n", len(rows))

	// 查找化学式列的索引
	formulaCol := ep.findFormulaColumn(rows[0])
	if formulaCol == -1 {
		return nil, 0, fmt.Errorf("未找到化学式列")
	}

	log.Printf("化学式列: 第 %d 列\n", formulaCol+1)

	// 处理数据行
	var emptyRows []int
	emptyCount := 0

	log.Printf("  开始扫描数据行...\n")

	for rowIndex := 1; rowIndex < len(rows); rowIndex++ {
		row := rows[rowIndex]

		// 确保行有足够的列
		if len(row) <= formulaCol {
			// 如果该行列数不足，也视为化学式为空
			emptyRows = append(emptyRows, rowIndex+1)
			emptyCount++
			continue
		}

		chemicalFormula := strings.TrimSpace(row[formulaCol])

		// 判断化学式是否为空
		if ep.isChemicalFormulaEmpty(chemicalFormula) {
			emptyRows = append(emptyRows, rowIndex+1)
			emptyCount++

			// 实时显示进度
			if emptyCount%500 == 0 {
				log.Printf("已找到 %d 个空化学式记录...\n", emptyCount)
			}
		}
	}

	log.Printf("工作表 %s: 找到 %d 个化学式为空的记录\n", sheetName, emptyCount)
	return emptyRows, emptyCount, nil
}

// findFormulaColumn 查找化学式列
func (ep *ExcelProcessor) findFormulaColumn(headers []string) int {
	for i, header := range headers {
		normalizedHeader := strings.ToLower(strings.TrimSpace(header))

		// 化学式列的可能名称
		formulaPatterns := []string{
			"化学式", "formula", "chemical formula", "chemicalformula",
			"分子式", "化学公式", "结构式", "chemical", "formula name",
			"化学结构", "分子结构", "结构式",
		}

		for _, pattern := range formulaPatterns {
			if strings.Contains(normalizedHeader, pattern) {
				log.Printf("识别化学式列: 第 %d 列 (%s)\n", i+1, header)
				return i
			}
		}
	}
	return -1
}

// isChemicalFormulaEmpty 判断化学式是否为空
func (ep *ExcelProcessor) isChemicalFormulaEmpty(formula string) bool {
	trimmed := strings.TrimSpace(formula)

	// 空字符串
	if trimmed == "" {
		return true
	}

	// 各种空值表示
	emptyPatterns := []string{
		"-", "--", "---", "----",
		"N/A", "NA", "n/a", "na", "N.A.",
		"NULL", "null", "nil",
		"未知", "不详", "无", "暂无", "未提供",
		"unknown", "none", "not available", "not provided",
		"待补充", "待定", "空缺", "缺",
		"TBD", "TBA", "待确认",
		"#N/A", "#REF!", "#VALUE!", "#NAME?",
	}

	for _, pattern := range emptyPatterns {
		if trimmed == pattern {
			return true
		}
	}

	// 检查是否只包含特殊字符或空格
	if strings.Trim(trimmed, ".-_/*\\|()[]{}<>~!@#$%^&* \t\n\r") == "" {
		return true
	}

	return false
}

// PrintResults 打印结果
func (ep *ExcelProcessor) PrintResults(emptyRows []int, totalCount int) {
	log.Printf("\n========== 统计结果 ==========\n")
	log.Printf("化学式为空的记录总数: %d\n", totalCount)
	log.Printf("空化学式所在行号列表:\n")
	log.Printf("==================================\n\n")

	// 分组显示行号（每行显示10个）
	for i := 0; i < len(emptyRows); i += 10 {
		end := i + 10
		if end > len(emptyRows) {
			end = len(emptyRows)
		}

		line := ""
		for j := i; j < end; j++ {
			line += fmt.Sprintf("%-6d", emptyRows[j])
		}
		log.Println(line)
	}

	log.Printf("\n==================================\n")
	log.Printf("总计: %d 个化学式为空的记录\n", totalCount)

	// 显示统计信息
	if totalCount > 0 {
		firstRow := emptyRows[0]
		lastRow := emptyRows[len(emptyRows)-1]
		log.Printf("行号范围: %d - %d\n", firstRow, lastRow)
		log.Printf("空记录占比: %.2f%%\n", float64(totalCount)/float64(lastRow)*100)
	}
}

// SaveToFile 保存结果到文件
func (ep *ExcelProcessor) SaveToFile(emptyRows []int, totalCount int, filename string) error {
	file, err := os.Create(filename)
	if err != nil {
		return err
	}
	defer file.Close()

	// 写入文件头
	file.WriteString("化学式为空的记录统计\n")
	file.WriteString("====================\n\n")
	file.WriteString(fmt.Sprintf("统计时间: %s\n", time.Now().Format("2006-01-02 15:04:05")))
	file.WriteString(fmt.Sprintf("源文件: %s\n", ep.FilePath))
	file.WriteString(fmt.Sprintf("总记录数: %d\n\n", totalCount))

	file.WriteString("空化学式所在行号:\n")
	file.WriteString("----------------\n")

	// 分组写入行号（每行10个）
	for i := 0; i < len(emptyRows); i += 10 {
		end := i + 10
		if end > len(emptyRows) {
			end = len(emptyRows)
		}

		line := ""
		for j := i; j < end; j++ {
			line += fmt.Sprintf("%-6d", emptyRows[j])
		}
		file.WriteString(line + "\n")
	}

	file.WriteString("\n====================\n")
	file.WriteString("统计完成!\n")

	return nil
}

// GenerateReport 生成详细报告
func (ep *ExcelProcessor) GenerateReport(emptyRows []int, totalCount int) {
	log.Printf("\n ========== 详细统计报告 ==========\n")
	log.Printf("文件名称: %s\n", ep.FilePath)
	log.Printf("处理时间: %s\n", time.Now().Format("2006-01-02 15:04:05"))
	log.Printf("空记录总数: %d\n", totalCount)

	// 空记录分布
	log.Printf("空记录分布:\n")
	log.Printf("最小行号: %d\n", emptyRows[0])
	log.Printf("最大行号: %d\n", emptyRows[len(emptyRows)-1])
	log.Printf("行号跨度: %d 行\n", emptyRows[len(emptyRows)-1]-emptyRows[0]+1)
}

// ParseExcel 解析excel
func ParseExcel(filePath string) map[int]string {
	// 初始化处理器
	processor := &ExcelProcessor{
		FilePath: filePath, // 替换为你的Excel文件路径
	}

	// 检查文件是否存在
	if _, err := os.Stat(processor.FilePath); os.IsNotExist(err) {
		log.Fatalf("文件不存在: %s", processor.FilePath)
	}

	log.Printf("开始解析Excel文件: %s\n", processor.FilePath)

	// 处理数据
	emptyRows, totalCount, err := processor.ProcessEmptyChemicalFormulas()
	if err != nil {
		log.Fatalf("处理失败: %v", err)
	}

	// 保存结果到文件
	outputFile := "../../docs/empty_formula_report.txt"
	cas, err := processor.GetCASByRowNumbers(emptyRows)
	if err != nil {
		log.Println("error!")
	}
	err = processor.SaveToFile(emptyRows, totalCount, outputFile)
	if err != nil {
		log.Printf("保存文件失败: %v", err)
	} else {
		log.Printf("结果已保存到: %s\n", outputFile)
	}
	return cas
}

// GetCASByRowNumbers 通过空行的行号获取cas号
func (ep *ExcelProcessor) GetCASByRowNumbers(rowNumbers []int) (map[int]string, error) {
	f, err := excelize.OpenFile(ep.FilePath)
	if err != nil {
		return nil, fmt.Errorf("打开文件失败: %v", err)
	}
	defer f.Close()

	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		return nil, fmt.Errorf("Excel 文件中没有工作表")
	}

	// 默认处理第一个工作表
	sheetName := sheets[0]
	return ep.getCASFromSheetBatch(f, sheetName, rowNumbers)
}

func (ep *ExcelProcessor) getCASFromSheetBatch(f *excelize.File, sheetName string, rowNumbers []int) (map[int]string, error) {
	// 获取所有行数据
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("读取行数据失败: %v", err)
	}

	if len(rows) == 0 {
		return nil, fmt.Errorf("工作表 %s 为空", sheetName)
	}

	// 查找CAS号列的索引
	casCol := ep.findCASColumn(rows[0])
	if casCol == -1 {
		return nil, fmt.Errorf("未找到CAS号列")
	}

	result := make(map[int]string)

	for _, rowNumber := range rowNumbers {
		// 检查行号是否有效
		if rowNumber < 1 || rowNumber > len(rows) {
			result[rowNumber] = fmt.Sprintf("错误: 行号超出范围")
			continue
		}

		// 获取指定行的数据
		row := rows[rowNumber-1]

		// 确保行有足够的列
		if len(row) <= casCol {
			result[rowNumber] = fmt.Sprintf("错误: 列数不足")
			continue
		}

		casNumber := strings.TrimSpace(row[casCol])
		result[rowNumber] = casNumber
	}
	return result, nil
}

func (ep *ExcelProcessor) findCASColumn(headers []string) int {
	for i, header := range headers {
		normalizedHeader := strings.ToLower(strings.TrimSpace(header))

		// CAS号列的可能名称
		casPatterns := []string{
			"cas", "cas号", "cas number", "cas no", "casno",
			"cas编号", "cas号码", "cas registry", "cas id",
			"卡斯", "卡斯号", "cas代码",
		}

		for _, pattern := range casPatterns {
			if strings.Contains(normalizedHeader, pattern) {
				log.Printf("识别CAS号列: 第 %d 列 (%s)\n", i+1, header)
				return i
			}
		}
	}
	return -1
}
