package main

import (
	"fmt"
	"io"
	"log"
	"net/http"

	"cas.mod/errorlog"
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

	for number, cas := range rowNumberAndCas {
		url := generateURL(cas)
		req, err := http.NewRequest("GET", url, nil)
		if err != nil {
			log.Println(err)
		}

		// 设置请求头
		req.Header.Set("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
		req.Header.Set("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8")
		req.Header.Set("Accept-Language", "zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3")
		req.Header.Set("Connection", "keep-alive")

		resp, err := http.DefaultClient.Do(req)
		if err != nil {
			fmt.Println("请求失败: ", err)
		}
		defer resp.Body.Close()

		// 检查状态码
		if resp.StatusCode != http.StatusOK {
			log.Printf("状态码错误: %d %s", resp.StatusCode, resp.Status)
			notExist++
			fmt.Println("404次数:", notExist)
			if err := errorlog.LogError(resp.StatusCode, url); err != nil {
				fmt.Println("写入URL到文件失败!")
			}
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
