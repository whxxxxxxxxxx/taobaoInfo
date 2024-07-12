package main

import (
	"bufio"
	"fmt"
	"os"
	"regexp"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	// 打开文件
	file, err := os.Open("tb信息.txt")
	if err != nil {
		fmt.Println("Error opening file:", err)
		return
	}
	defer file.Close()

	// 读取文件内容
	var data string
	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		data += scanner.Text() + "\n"
	}

	if err := scanner.Err(); err != nil {
		fmt.Println("Error reading file:", err)
		return
	}

	// 分割多个订单的信息
	orderSeparator := regexp.MustCompile(`(?m)^订单号：`)
	orders := orderSeparator.Split(data, -1)

	// 创建一个新的 Excel 文件
	f := excelize.NewFile()

	// 设置表头
	headers := []string{"收件人姓名（必填）", "收件人手机（二选一）", "收件人电话（二选一）", "收件人地址（必填）",
		"寄件人姓名", "寄件人手机（二选一）", "寄件人电话（二选一）", "寄件人地址", "物品类型（最多4个字）", "包裹备注", "订单编号", "包裹重量", "包裹体积"}
	for i, header := range headers {
		cell := fmt.Sprintf("%s1", string(rune('A'+i)))
		f.SetCellValue("Sheet1", cell, header)
	}

	// 处理每个订单
	row := 2
	for _, order := range orders {
		if strings.TrimSpace(order) == "" {
			continue
		}

		// 提取地址和联系人信息
		reAddress := regexp.MustCompile(`--\n(.*)\n`)
		addressInfo := reAddress.FindStringSubmatch(order)
		if addressInfo[1] == "已修改地址" {
			reAddress = regexp.MustCompile(`--\n.*\n(.*)\n`)
			addressInfo = reAddress.FindStringSubmatch(order)
		}
		if len(addressInfo) > 1 {
			addressLine := strings.TrimSpace(addressInfo[1])

			// 提取收件人姓名和电话
			reContact := regexp.MustCompile(`^(.*)，(\d+)-(\d+)，(.*)`)
			contactInfo := reContact.FindStringSubmatch(addressLine)
			if len(contactInfo) > 4 {
				name := contactInfo[1] + "[" + contactInfo[3] + "]"
				phone := contactInfo[2]
				address := contactInfo[4]

				// 提取商家编码和下一行信息
				reMerchant := regexp.MustCompile(`商家编码：(.*)`)
				merchantInfo := reMerchant.FindStringSubmatch(order)
				merchantCodeLine := "未找到"
				if len(merchantInfo) > 1 {
					merchantCodeLine = strings.TrimSpace(merchantInfo[1])
				}

				// 设置单元格内容
				f.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), name)
				f.SetCellValue("Sheet1", fmt.Sprintf("B%d", row), phone)
				f.SetCellValue("Sheet1", fmt.Sprintf("D%d", row), address)
				f.SetCellValue("Sheet1", fmt.Sprintf("J%d", row), merchantCodeLine)

				row++
				fmt.Println(row)
			} else {
				reContact = regexp.MustCompile(`^(.*)，(\d+)，(.*)`)
				contactInfo = reContact.FindStringSubmatch(addressLine)
				name := contactInfo[1]
				phone := contactInfo[2]
				address := contactInfo[3]

				// 提取商家编码和下一行信息
				reMerchant := regexp.MustCompile(`商家编码：(.*)`)
				merchantInfos := reMerchant.FindAllStringSubmatch(order, -1)
				merchantCodeLine := "未找到"
				for _, merchantInfo := range merchantInfos {
					if len(merchantInfo) > 1 {
						if merchantCodeLine == "未找到" {
							merchantCodeLine = strings.TrimSpace(merchantInfo[1])
						} else {
							merchantCodeLine = merchantCodeLine + " " + strings.TrimSpace(merchantInfo[1])
						}

					}
				}

				// 设置单元格内容
				f.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), name)
				f.SetCellValue("Sheet1", fmt.Sprintf("B%d", row), phone)
				f.SetCellValue("Sheet1", fmt.Sprintf("D%d", row), address)
				f.SetCellValue("Sheet1", fmt.Sprintf("J%d", row), merchantCodeLine)

				row++
			}
		}
	}
	timestamp := time.Now().Format("20060102_150405")
	filename := fmt.Sprintf("orders_%s.xlsx", timestamp)
	// 保存 Excel 文件
	if err := f.SaveAs(filename); err != nil {
		fmt.Println("Error saving file:", err)
	}
}
