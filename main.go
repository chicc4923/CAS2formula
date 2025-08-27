package main

import (
	"fmt"
	"io"
	"log"
	"net/http"

	"cas.mod/excel"
	parsehtml "cas.mod/parseHTML"
	"golang.org/x/text/encoding/simplifiedchinese"
	"golang.org/x/text/transform"
)

func generateURL(cas string) string {
	return fmt.Sprintf("http://www.ichemistry.cn/chemistry/" + cas + ".htm")
}

func main() {
	filePath := "ReagentModules.xlsx"
	rowNumberAndCas := excel.ParseExcel(filePath)
	var notExist int
	// excel.WriteTestToFormulaCells(filePath, emptyRows)

	for number, cas := range rowNumberAndCas {
		url := generateURL(cas)
		// 发送 GET 请求
		resp, err := http.Get(url)
		if err != nil {
			fmt.Println("请求失败: ", err)
		}
		defer resp.Body.Close()

		// 检查状态码
		if resp.StatusCode != http.StatusOK {
			log.Printf("状态码错误: %d %s", resp.StatusCode, resp.Status)
			notExist++
			fmt.Println("404次数:", notExist)
			continue
		}

		// 创建GBK解码器Reader
		reader := transform.NewReader(resp.Body, simplifiedchinese.GBK.NewDecoder())

		// 读取解码后的内容
		body, err := io.ReadAll(reader)
		if err != nil {
			log.Fatal("读取响应失败: ", err)
		}

		// 解析HTML
		parsehtml.ParseHTML(string(body), 1, number)
	}
}
