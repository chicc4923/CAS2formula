package errorlog

import (
	"bufio"
	"fmt"
	"os"
	"time"
)

// logError 记录错误状态码和URL到文件
func LogError(statusCode int, url string) error {
	// 打开文件（如果不存在则创建，以追加模式打开）
	file, err := os.OpenFile("error_log.txt", os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
	if err != nil {
		return fmt.Errorf("无法打开错误日志文件: %v", err)
	}
	defer file.Close()

	// 获取当前时间
	currentTime := time.Now().Format("2006-01-02 15:04:05")

	// 创建写入器
	writer := bufio.NewWriter(file)

	// 格式化日志内容：时间戳 | 状态码 | URL
	logEntry := fmt.Sprintf("%s | %d | %s\n", currentTime, statusCode, url)

	// 写入文件
	_, err = writer.WriteString(logEntry)
	if err != nil {
		return fmt.Errorf("写入错误日志失败: %v", err)
	}

	// 确保内容刷新到磁盘
	err = writer.Flush()
	if err != nil {
		return fmt.Errorf("刷新错误日志失败: %v", err)
	}

	return nil
}
