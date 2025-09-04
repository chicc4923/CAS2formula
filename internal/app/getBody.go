package app

import (
	"fmt"
	"log"
	"net/http"
	"strings"

	"github.com/PuerkitoBio/goquery"
)

// ChemicalInfo 化学信息结构体
type ChemicalInfo struct {
	CASNumber       string // CAS号
	ChineseName     string // 中文名
	EnglishName     string // 英文名
	ChemicalFormula string // 化学式
	StructureImage  string // 结构式图片URL
}

// GetChemicalInfo 根据CAS号获取化学信息
func GetChemicalInfo(casNumber string) (*ChemicalInfo, error) {
	// 构建URL
	url := fmt.Sprintf("http://search.ichemistry.cn/?keys=%s&onlymy=0&types=2&tz=1", casNumber)

	// 发送HTTP GET请求
	resp, err := http.Get(url)
	if err != nil {
		return nil, fmt.Errorf("HTTP请求失败: %v", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != 200 {
		return nil, fmt.Errorf("HTTP状态码错误: %d %s", resp.StatusCode, resp.Status)
	}

	// 使用goquery解析HTML
	doc, err := goquery.NewDocumentFromReader(resp.Body)
	if err != nil {
		return nil, fmt.Errorf("解析HTML失败: %v", err)
	}

	// 创建化学信息对象
	info := &ChemicalInfo{CASNumber: casNumber}
	found := false

	// 查找包含化学信息的表格
	doc.Find("table#container-right tr").Each(func(i int, s *goquery.Selection) {
		// 跳过表头行
		if i == 0 {
			return
		}

		// 检查是否包含目标CAS号
		if strings.Contains(s.Text(), casNumber) {
			found = true
			s.Find("td").Each(func(j int, td *goquery.Selection) {
				text := strings.TrimSpace(td.Text())
				switch j {
				case 1: // 中文名
					info.ChineseName = text
				case 2: // 英文名
					info.EnglishName = text
				case 3: // 结构式图片
					if img := td.Find("img"); img.Length() > 0 {
						if src, exists := img.Attr("src"); exists {
							info.StructureImage = src
						}
					}
				case 4: // 化学式
					info.ChemicalFormula = text
				}
			})
		}
	})

	if !found {
		return nil, fmt.Errorf("未找到CAS号 %s 的信息", casNumber)
	}

	return info, nil
}

// GetChemicalInfoWithCustomURL 使用自定义URL获取化学信息
func GetChemicalInfoWithCustomURL(url string) (*ChemicalInfo, error) {
	// 发送HTTP GET请求
	resp, err := http.Get(url)
	if err != nil {
		return nil, fmt.Errorf("HTTP请求失败: %v", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != 200 {
		return nil, fmt.Errorf("HTTP状态码错误: %d %s", resp.StatusCode, resp.Status)
	}

	// 使用goquery解析HTML
	doc, err := goquery.NewDocumentFromReader(resp.Body)
	if err != nil {
		return nil, fmt.Errorf("解析HTML失败: %v", err)
	}

	info := &ChemicalInfo{}
	found := false

	// 查找包含化学信息的表格
	doc.Find("table#container-right tr").Each(func(i int, s *goquery.Selection) {
		// 跳过表头行
		if i == 0 {
			return
		}

		// 从第一列提取CAS号
		s.Find("td").Each(func(j int, td *goquery.Selection) {
			if j == 0 {
				cas := strings.TrimSpace(td.Text())
				if cas != "" {
					info.CASNumber = cas
					found = true
				}
			}
		})

		if found {
			s.Find("td").Each(func(j int, td *goquery.Selection) {
				text := strings.TrimSpace(td.Text())
				switch j {
				case 1: // 中文名
					info.ChineseName = text
				case 2: // 英文名
					info.EnglishName = text
				case 3: // 结构式图片
					if img := td.Find("img"); img.Length() > 0 {
						if src, exists := img.Attr("src"); exists {
							info.StructureImage = src
						}
					}
				case 4: // 化学式
					info.ChemicalFormula = text
				}
			})
		}
	})

	if !found {
		return nil, fmt.Errorf("未找到化学信息")
	}

	return info, nil
}

// PrintChemicalInfo 打印化学信息
func PrintChemicalInfo(info *ChemicalInfo) {
	log.Println("=" + strings.Repeat("=", 50))
	log.Printf("CAS号: %s\n", info.CASNumber)
	log.Printf("中文名: %s\n", info.ChineseName)
	log.Printf("英文名: %s\n", info.EnglishName)
	log.Printf("化学式: %s\n", info.ChemicalFormula)
	if info.StructureImage != "" {
		log.Printf("结构式图片: %s\n", info.StructureImage)
	}
	log.Println("=" + strings.Repeat("=", 50))
}

// GetChemicalFormulaOnly 仅获取化学式
func GetChemicalFormulaOnly(casNumber string) (string, error) {
	info, err := GetChemicalInfo(casNumber)
	if err != nil {
		return "", err
	}
	return info.ChemicalFormula, nil
}

// BatchGetChemicalInfo 批量获取化学信息
func BatchGetChemicalInfo(casNumbers []string) (map[string]*ChemicalInfo, []error) {
	results := make(map[string]*ChemicalInfo)
	var errors []error

	for _, casNumber := range casNumbers {
		info, err := GetChemicalInfo(casNumber)
		if err != nil {
			errors = append(errors, fmt.Errorf("CAS %s: %v", casNumber, err))
			continue
		}
		results[casNumber] = info
	}

	return results, errors
}
