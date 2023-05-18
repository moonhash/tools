package main

import (
	"fmt"
	"github.com/nguyenthenguyen/docx"
	"github.com/sirupsen/logrus"
	"github.com/xuri/excelize/v2"
	"gopkg.in/natefinch/lumberjack.v2"
	"os"
	"strconv"
	"strings"
)

func main() {
	logrus.SetReportCaller(true)
	dir, err := os.Getwd()
	if err != nil {
		logrus.Printf("获取当前目录路劲失败, err: %v", err)
	}
	logrus.SetOutput(&lumberjack.Logger{
		Filename:   dir + "/run.log",
		MaxSize:    1000, // megabytes
		MaxBackups: 6,
		MaxAge:     15, //days
		LocalTime:  true,
	})
	//f, err := excelize.OpenFile("./填写信息汇总.xlsx")
	//if err != nil {
	//	fmt.Println(err)
	//	return
	//}
	//defer func() {
	//	// Close the spreadsheet.
	//	if err := f.Close(); err != nil {
	//		fmt.Println(err)
	//	}
	//}()
	//sheetName := "Sheet1"
	//templateName := "template.docx"
	//rows, err := f.GetRows(sheetName)
	//if err != nil {
	//	fmt.Printf("无法打开%v表格", sheetName)
	//	return
	//}
	//if len(rows) <= 1 {
	//	fmt.Printf("填写信息汇总xlsx无内容")
	//	return
	//}
	//fields := make(map[string]int, 12)
	//// fmt.Println(rows)
	//for i, v := range rows[0] {
	//	fields[strings.TrimSpace(v)] = i
	//}
	//for i, v := range rows {
	//	if i == 0 {
	//		continue
	//	}
	//	// replaceMap is a key-value map whereas the keys
	//	// represent the placeholders without the delimiters
	//	letterNo := string(byte(65 + fields["签名图片"]))
	//	picture, err := f.GetPictures(sheetName, letterNo+strconv.Itoa(i+1))
	//	if err != nil {
	//		fmt.Printf("获取前面图片失败, err: %v", err.Error())
	//		return
	//	}
	//	if len(picture) == 0 {
	//		fmt.Printf("签名图片为空")
	//		return
	//	}
	//	pictureName := v[fields["法人名字"]] + picture[0].Extension
	//	err = os.WriteFile(pictureName, picture[0].File, 0755)
	//	if err != nil {
	//		fmt.Printf("保存签名图片失败, err: %v", err.Error())
	//		return
	//	}
	//	replaceMap := docx.PlaceholderMap{
	//		"key-with-name":      v[fields["法人名字"]],
	//		"key-with-naciona":   v[fields["国籍"]],
	//		"key-with-pasaporte": v[fields["身份证号码"]],
	//		"key-with-company":   v[fields["公司名"]],
	//		"key-with-addr":      v[fields["公司地址"]],
	//		"key-with-postal":    v[fields["邮编"]],
	//		"key-with-provincia": v[fields["省份"]],
	//		"key-with-country":   v[fields["国家"]],
	//		"key-with-day":       v[fields["日"]],
	//		"key-with-month":     v[fields["月"]],
	//		"key-with-year":      v[fields["年"]],
	//	}
	//
	//	// read and parse the template docx
	//	doc, err := docx.Open(templateName)
	//	if err != nil {
	//		fmt.Printf("打开模板docx文件失败, err: %v", err.Error())
	//		return
	//	}
	//
	//	// replace the keys with values from replaceMap
	//	err = doc.ReplaceAll(replaceMap)
	//	if err != nil {
	//		fmt.Printf("替换模板文件内容失败, err: %v", err.Error())
	//		return
	//	}
	//
	//	// write out a new file
	//	err = doc.WriteToFile(strings.TrimSpace(v[fields["文件命名"]]) + ".docx")
	//	if err != nil {
	//		fmt.Printf("保存为新文件失败, err: %v", err.Error())
	//		return
	//	}
	//	r, err := docx2.ReadDocxFile(templateName)
	//	if err != nil {
	//		fmt.Printf("docx2打开模板文件失败, err: %v", err.Error())
	//		return
	//	}
	//	docxNew := r.Editable()
	//	err = docxNew.Replace("{key-with-fdo}", "123", -1)
	//	if err != nil {
	//		fmt.Printf("docx2更换失败, err: %v", err.Error())
	//		return
	//	}
	//	err = docxNew.WriteToFile("111.docx")
	//	if err != nil {
	//		fmt.Printf("docx2保存为新文件失败, err: %v", err.Error())
	//		return
	//	}
	//	return
	//}
	feiDuDu(dir)
	// test()
}

func feiDuDu(dir string) {
	f, err := excelize.OpenFile("./填写信息汇总.xlsx")
	if err != nil {
		logrus.Println(err)
		return
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			logrus.Println(err)
		}
	}()
	sheetName := "Sheet1"
	templateName := dir + "/template.docx"
	fmt.Println(templateName)
	rows, err := f.GetRows(sheetName)
	if err != nil {
		logrus.Printf("无法打开%v表格", sheetName)
		return
	}
	if len(rows) <= 1 {
		logrus.Printf("填写信息汇总xlsx无内容")
		return
	}
	fields := make(map[string]int, 12)
	// fmt.Println(rows)
	for i, v := range rows[0] {
		fields[strings.TrimSpace(v)] = i
	}
	r, err := docx.ReadDocxFile(templateName)
	if err != nil {
		logrus.Printf("docx打开模板文件失败, err: %v", err.Error())
		return
	}
	defer r.Close()
	for i, v := range rows {
		if i == 0 {
			continue
		}
		// replaceMap is a key-value map whereas the keys
		// represent the placeholders without the delimiters
		letterNo := string(byte(65 + fields["签名图片"]))
		picture, err := f.GetPictures(sheetName, letterNo+strconv.Itoa(i+1))
		if err != nil {
			logrus.Printf("获取前面图片失败, err: %v", err.Error())
			return
		}
		if len(picture) == 0 {
			logrus.Printf("签名图片为空")
			return
		}
		pictureName := v[fields["法人名字"]] + picture[0].Extension
		err = os.WriteFile(pictureName, picture[0].File, 0755)
		if err != nil {
			logrus.Printf("保存签名图片失败, err: %v", err.Error())
			return
		}
		replaces := map[string]string{
			"{key-with-name}":      v[fields["法人名字"]],
			"{key-with-naciona}":   v[fields["国籍"]],
			"{key-with-pasaporte}": v[fields["身份证号码"]],
			"{key-with-company}":   v[fields["公司名"]],
			"{key-with-addr}":      v[fields["公司地址"]],
			"{key-with-postal}":    v[fields["邮编"]],
			"{key-with-provincia}": v[fields["省份"]],
			"{key-with-country}":   v[fields["国家"]],
			"{key-with-day}":       v[fields["日"]],
			"{key-with-month}":     v[fields["月"]],
			"{key-with-year}":      v[fields["年"]],
		}
		docxNew := r.Editable()
		for k, s := range replaces {
			err = docxNew.Replace(k, s, -1)
			if err != nil {
				logrus.Printf("docx更换失败, err: %v", err.Error())
				return
			}
		}
		err = docxNew.ReplaceImage("word/media/image1.png", pictureName)
		if err != nil {
			logrus.Printf("docx更换图片失败, err: %v", err.Error())
			return
		}
		err = docxNew.WriteToFile(strings.TrimSpace(v[fields["文件命名"]]) + ".docx")
		if err != nil {
			logrus.Printf("docx保存为新文件失败, err: %v", err.Error())
			return
		}
		os.Remove(pictureName)
	}
}

func test() {
	dir, _ := os.Getwd()
	logrus.Println(dir)
	f, _ := os.Executable()
	logrus.Println(f)

}
