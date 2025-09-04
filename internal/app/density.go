package app

import (
	"fmt"
	"log"
	"strings"

	"github.com/PuerkitoBio/goquery"
)

// Density 获取密度函数
func Density(body string) {
	// 使用goquery解析HTML
	doc, err := goquery.NewDocumentFromReader(strings.NewReader(body))
	if err != nil {
		log.Fatal("解析HTML失败:", err)
	}

	// 查找密度值
	density := ""

	// 方法1：通过表格结构定位
	doc.Find("table#baseTbl tr").Each(func(i int, s *goquery.Selection) {
		// 查找包含"密度"的行
		if strings.Contains(s.Text(), "密度") {
			// 获取密度值所在的单元格
			s.Find("td").Each(func(j int, td *goquery.Selection) {
				if j == 1 { // 密度值通常在第二列
					density = strings.TrimSpace(td.Text())
				}
			})
		}
	})

	// 方法2：通过ID直接定位（更可靠）
	if density == "" {
		doc.Find("#wuHuaDiv table tr").Each(func(i int, s *goquery.Selection) {
			th := s.Find("th").Text()
			if strings.Contains(th, "密度") {
				density = strings.TrimSpace(s.Find("td").Text())
			}
		})
	}

	if density != "" {
		fmt.Printf("密度值: %s\n", density)
	} else {
		fmt.Println("未找到密度信息")

		// 调试：打印所有表格内容
		doc.Find("table").Each(func(i int, s *goquery.Selection) {
			id, _ := s.Attr("id")
			fmt.Printf("表格 %d (ID: %s):\n", i+1, id)
			fmt.Println(s.Text())
			fmt.Println("------")
		})
	}
}
