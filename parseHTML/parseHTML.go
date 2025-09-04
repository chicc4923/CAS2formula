package parsehtml

import (
	"log"
	"strings"

	"cas.mod/excel"
	"github.com/PuerkitoBio/goquery"
)

func ParseHTML(htmlContent string, unknown int, number int) {
	doc, err := goquery.NewDocumentFromReader(strings.NewReader(htmlContent))
	if err != nil {
		log.Fatal("解析HTML错误: ", err)
	}

	// 查找分子式
	var molecularFormula string
	doc.Find("table.ChemicalInfo tr").Each(func(i int, s *goquery.Selection) {
		// 查找包含"分子式"文本的单元格
		if strings.Contains(s.Text(), "分子式") {
			// 获取相邻单元格的内容
			s.Find("td").Each(func(j int, td *goquery.Selection) {
				if j == 1 { // 第二个td包含分子式
					molecularFormula = td.Text()
					molecularFormula = strings.TrimSpace(molecularFormula)
				}
			})
		}
	})

	if molecularFormula != "" {
		log.Printf("number: %v", number)
		log.Printf("找到分子式: %s\n", molecularFormula)
		excel.WriteToCell("ReagentModules.xlsx", "Sheet1", "化学式", number, molecularFormula)
	} else {
		unknown++
		// 尝试查看实际内容（调试用）
		log.Println("未找到分子式。可能的表格行:")
		doc.Find("table.ChemicalInfo tr").Each(func(i int, s *goquery.Selection) {
			log.Printf("行 %d: %s\n", i, s.Text())
		})
		log.Println("尝试使用CSS选择器定位分子式")

		// 直接使用CSS选择器定位分子式行
		doc.Find("tr:has(td.ltd:contains('分子式'))").Each(func(i int, s *goquery.Selection) {
			s.Find("td").Each(func(j int, td *goquery.Selection) {
				if j == 1 {
					molecularFormula = td.Text()
					molecularFormula = strings.TrimSpace(molecularFormula)
					log.Printf("通过CSS选择器找到分子式: %s\n", molecularFormula)
				}
			})
		})
	}
}
