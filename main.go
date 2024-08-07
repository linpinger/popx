package main

import (
	"bytes"
	"flag"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"mime"
	"net/mail"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	msg "github.com/emersion/go-message"
	eml "github.com/emersion/go-message/mail"
	_ "github.com/emersion/go-message/charset"
	"github.com/knadh/go-pop3"
)

var (
	VerStr  string = "2024-08-07.10"
	XmlOnlinePATH  string = "00_eml_online.xml"
	XmlOfflinePATH string = "00_eml_offline.xml"
)


func main() { // 下载邮件

	// 命令行参数
	flag.Usage = func() {
		fmt.Println("# Version:", VerStr)
		fmt.Println("# Usage:", os.Args[0], "[args]")
		flag.PrintDefaults()
		fmt.Println("# Example:")
		fmt.Println("  [set|export] eSvrPOP=pop.163.com:995:1")
		fmt.Println("  [set|export] eUP=xx@163.com:hello")
		fmt.Println("  ", os.Args[0], "-l 9                 列出最近9封邮件标题")
		fmt.Println("  ", os.Args[0], "[-b] [-a] -d 1       下载序列号为1的邮件并释放[正文][附件]")
		fmt.Println("  ", os.Args[0], "-rm 1                删除序列号为1的邮件")
		fmt.Println("  ", os.Args[0], "-da 9                倒序下载最近9封邮件")
		fmt.Println("  ", os.Args[0], "-n                   离线: 重命名.里的eml并将信息写入:"+XmlOfflinePATH)
		fmt.Println("  ", os.Args[0], "[-b] [-a] -f xx.eml  离线: 释放xx.eml中的[正文][附件]")
		os.Exit(0)
	}
	bRenameEmls := false
	flag.BoolVar(&bRenameEmls, "n", bRenameEmls, "离线: 重命名.里的eml并将信息写入:"+XmlOfflinePATH)
	var downCount int
	flag.IntVar(&downCount, "da", -1, "倒序下载最近n封邮件")
	var listCount int
	flag.IntVar(&listCount, "l", -1, "倒序列出最近n封邮件的标题")
	var downIDX int
	flag.IntVar(&downIDX, "d", -1, "下载序列号为N的邮件并释放附件")
	var deleteIDX int
	flag.IntVar(&deleteIDX, "rm", -1, "删除序列号为N的邮件")

	emlPath := "xx.eml"
	flag.StringVar(&emlPath, "f", emlPath, "离线: 释放"+emlPath+"中的[正文b][附件a]")
	bExtractBody := false
	flag.BoolVar(&bExtractBody, "b", bExtractBody, "释放eml中的正文")
	bExtractAttachMent := false
	flag.BoolVar(&bExtractAttachMent, "a", bExtractAttachMent, "释放eml中的附件")

	flag.Parse() // 处理参数

	if bRenameEmls {
		renameEmlFiles()
		os.Exit(0)
	}
	if "xx.eml" != emlPath {
		fEML, _ := os.Open(emlPath) // 读取eml
		emlParser(fEML, bExtractBody, bExtractAttachMent) // 释放 正文 附件
		os.Exit(0)
	}

	NowMode := "blank" // NowMode: list, downAll, downOne, deleteOne
	NID := 1
	if downCount != -1 {
		NowMode = "downAll"
		NID = downCount
	}
	if listCount != -1 {
		NowMode = "list"
		NID = listCount
	}
	if downIDX != -1 {
		NowMode = "downOne"
		if downIDX < 1 {
			fmt.Println("- Error:", downIDX, "< 1")
			os.Exit(1)
		}
	}
	if deleteIDX != -1 {
		NowMode = "deleteOne"
		if deleteIDX < 1 {
			fmt.Println("- Error:", deleteIDX, "< 1")
			os.Exit(1)
		}
	}
	if NID < 1 {
		fmt.Println("- Error:", NID, "< 1")
		os.Exit(1)
	}

	if "blank" == NowMode { // 参数不对或无参数
		flag.Usage()
		os.Exit(1)
	}
	// 下面是在线模式

	// 环境变量
	envSV := os.Getenv("eSvrPOP")
	envUP := os.Getenv("eUP")
	if "" == envUP {
		envUP = "hello@163.com:helloworld"
	}
	aUP := strings.Split(envUP, ":")
	if "" == aUP[0] || "" == aUP[1] {
		return
	}
	if "" == envSV {
		envSV = "pop.ym.163.com:995:1"
	}
	aEml := strings.Split(envSV, ":")
	bTLS := false
	if "" == aEml[0] || "" == aEml[1] || "" == aEml[2] {
		envSV = "pop.ym.163.com:995:1"
		aEml = strings.Split(envSV, ":")
	}
	if "1" == aEml[2] {
		bTLS = true
	}
	nPort, err := strconv.Atoi(aEml[1])
	if err != nil {
		log.Fatal(err)
	}

	var XmlFile *os.File
	XmlStr := ""
	if "downAll" == NowMode {
		// 读取 xml 里面包含的 eml 信息
		xmlBytes, _ := os.ReadFile(XmlOnlinePATH)
		XmlStr = string(xmlBytes)
		fXmlTmp, err := os.OpenFile(XmlOnlinePATH, os.O_RDWR|os.O_CREATE|os.O_APPEND, 0644)
		if err != nil {
			log.Fatal(err)
		}
		XmlFile = fXmlTmp
		// defer XmlFile.Close()
	}

	// 连接
	p := pop3.New(pop3.Opt{
		Host:       aEml[0],
		Port:       nPort,
		TLSEnabled: bTLS,
	})
	// 110, 995

	c, err := p.NewConn()
	if err != nil {
		log.Fatal(err)
	}
	defer c.Quit()

	if err := c.Auth(aUP[0], aUP[1]); err != nil {
		log.Fatal(err)
	}
	fmt.Printf("# %s @ %s:%d  isTLS:%t\n", aUP[0], aEml[0], nPort, bTLS)

	count, size, _ := c.Stat()
	fmt.Println("# total messages =", count, ", size =", size)

	if "downOne" == NowMode { // downIDX
		if downIDX > count {
			downIDX = count
		}
		if 0 != count {
			emlPath := fmt.Sprintf("%d.eml", downIDX)
			buf, _ := c.RetrRaw(downIDX)
			saveEml(buf, emlPath) // 保存eml: ret: writeLen

			if bExtractBody || bExtractAttachMent {
				fEML, _ := os.Open(emlPath) // 读取eml
				emlParser(fEML, bExtractBody, bExtractAttachMent) // 释放 正文 附件
			} else {
				fmt.Println("+", downIDX, ":", emlPath)
			}
		}
	}

	if "deleteOne" == NowMode { // deleteIDX
		if deleteIDX > count {
			deleteIDX = count
		}
		if 0 != count {
			m, _ := c.Top(deleteIDX, 0) // 取头
			nowSubject, _ := m.Header.Text("subject")
			fmt.Println("- 删除 :", deleteIDX, ":", nowSubject)

			err = c.Dele(deleteIDX)
			if err != nil {
				fmt.Println("- 错误:", err)
			}
		}
	}

	if "downAll" == NowMode || "list" == NowMode {
		if NID > count {
			NID = 1
		} else {
			NID = count - NID + 1 // 取最后NID封邮件
		}
		if count == 0 {
			NID = 9
		}

		for id := count; id >= NID; id-- {
			bDownEML := true // 通过判断xml中是否包含UID来确定是否下载eml
			nowUID := "" // 当前UID
			aUID, _ := c.Uidl(id)
			if 1 == len(aUID) {
				nowUID = aUID[0].UID
			}
			if len(nowUID) > 1 && strings.Contains(XmlStr, nowUID) {
				bDownEML = false
			}

			m, _ := c.Top(id, 0) // 取头
			ei := NewEmlInfoFromMessageHeader(m.Header).SetUID(nowUID)

			fmt.Println("|", id, ":", ei.GetLocalTime(), ":", ei.Subject, ":", ei.From, ":", nowUID)

			if "downAll" == NowMode && bDownEML {
				buf, _ := c.RetrRaw(id)

				emlName0 := fmt.Sprintf("00_tmp_%d.eml", id) // 初始名称: 后面会更名
				writeLen := saveEml(buf, emlName0) // 保存eml
				ei.SetLength(writeLen)
				emlName1 := ei.GetFileName1()
				os.Rename(emlName0, emlName1) // 第1次重命名
				emlName2 := ei.GetFileName2()
				os.Rename(emlName1, emlName2) // 第2次重命名
				fmt.Printf("+ %d : %s\n", id, emlName2)

				if _, err := XmlFile.WriteString(ei.GetXMLLine()); err != nil {
					XmlFile.Close()
					log.Fatal(err)
				}
			} // downAll
		}

		if "downAll" == NowMode {
			if err := XmlFile.Close(); err != nil {
				log.Fatal(err)
			}
		}
	} // "downAll" || "list"

	fmt.Println("# total messages =", count, ", size =", size)
}

func niceFileName(iStr string) string {
	iStr = strings.Replace(iStr, "*", "※", -1)
	iStr = strings.Replace(iStr, "\\", "﹨", -1)
	iStr = strings.Replace(iStr, "|", "｜", -1)
	iStr = strings.Replace(iStr, ":", "︰", -1)
	iStr = strings.Replace(iStr, "\"", "¨", -1)
	iStr = strings.Replace(iStr, "<", "〈", -1)
	iStr = strings.Replace(iStr, ">", "〉", -1)
	iStr = strings.Replace(iStr, "/", "／", -1)
	iStr = strings.Replace(iStr, "?", "？", -1)
	return iStr
}

func stupidDate(sDate string) (time.Time, error) { // 中文月份替换为英文 "16 九月 2023 12:38:01 +/-TZ" -> "16 Sep 2023 12:38:01 +0800"
	if strings.Contains(sDate, "+/-TZ") {
		sDate = strings.Replace(sDate, "+/-TZ", "+0800", -1)

		sDate = strings.Replace(sDate, "十一月", "Nov", -1)
		sDate = strings.Replace(sDate, "11月", "Nov", -1)
		sDate = strings.Replace(sDate, "十二月", "Dec", -1)
		sDate = strings.Replace(sDate, "12月", "Dec", -1)

		sDate = strings.Replace(sDate, "一月", "Jan", -1)
		sDate = strings.Replace(sDate, "1月", "Jan", -1)
		sDate = strings.Replace(sDate, "二月", "Feb", -1)
		sDate = strings.Replace(sDate, "2月", "Feb", -1)
		sDate = strings.Replace(sDate, "三月", "Mar", -1)
		sDate = strings.Replace(sDate, "3月", "Mar", -1)
		sDate = strings.Replace(sDate, "四月", "Apr", -1)
		sDate = strings.Replace(sDate, "4月", "Apr", -1)
		sDate = strings.Replace(sDate, "五月", "May", -1)
		sDate = strings.Replace(sDate, "5月", "May", -1)
		sDate = strings.Replace(sDate, "六月", "Jun", -1)
		sDate = strings.Replace(sDate, "6月", "Jun", -1)
		sDate = strings.Replace(sDate, "七月", "Jul", -1)
		sDate = strings.Replace(sDate, "7月", "Jul", -1)
		sDate = strings.Replace(sDate, "八月", "Aug", -1)
		sDate = strings.Replace(sDate, "8月", "Aug", -1)
		sDate = strings.Replace(sDate, "九月", "Sep", -1)
		sDate = strings.Replace(sDate, "9月", "Sep", -1)
		sDate = strings.Replace(sDate, "十月", "Oct", -1)
		sDate = strings.Replace(sDate, "10月", "Oct", -1)
	}
	return mail.ParseDate(sDate)
}

// 这个函数改自: github.com/emersion/go-message/mail/header.go
func emlDate(strDate string) (time.Time, error) {
	if strDate == "" {
		return time.Time{}, nil
	}
//	return mail.ParseDate(strDate)
	return stupidDate(strDate)
}

// 读取eml，获取信息，释放正文和附件
func emlParser(r io.Reader, bExtractBody bool, bExtractAttachMent bool) {
	mr, _ := eml.CreateReader(r)

//	iFrom, _ := mr.Header.AddressList("From") // []*Address
	iFrom, _ := mr.Header.Text("From")
	fmt.Println("# From:", iFrom)

	sTo, _ := mr.Header.Text("To")
	fmt.Println("# To:", sTo)

	sCC, _ := mr.Header.Text("CC")
	if len(sCC) > 0 {
		fmt.Println("# CC:", sCC)
	}

	sBCC, _ := mr.Header.Text("Bcc")
	if len(sBCC) > 0 {
		fmt.Println("# BCC:", sBCC)
	}

	iDate, _    := emlDate(mr.Header.Get("date"))
	iSubject, _ := mr.Header.Text("subject")

	fmt.Println("# Date:", iDate)
	fmt.Println("# LocalDate:", iDate.Local())
	fmt.Println("# TimeStamp:", iDate.Unix())
	fmt.Println("# Subject:", iSubject)

	fmt.Println("---------------------")

	inlineCount := 0
	for {
		p, err := mr.NextPart() // *Part
		if err == io.EOF {
			break
		} else if err != nil {
			fmt.Println(err)
		}

		switch h := p.Header.(type) {
		case *eml.InlineHeader:
			inlineCount = 1 + inlineCount

			oExt := ".txt"
			ct, ctpar, _ := h.ContentType() // (t string, params map[string]string, err error)
			if strings.Contains(ct, "html") {
				oExt = ".html"
			}
			if strings.Contains(ct, "image/") {
				picName, err := getInlineFilename(h) // 自己写的
				if err != nil {
					oExt = ".jpg"
					aExt, _ := mime.ExtensionsByType(ct)
					if len(aExt) > 0 {
						oExt = aExt[0]
					}
				}
				if "" != picName && bExtractBody {
					fmt.Printf("# %d.%s : %s : %v\n", inlineCount, picName, ct, ctpar)
					b, _ := ioutil.ReadAll(p.Body)
					os.WriteFile(fmt.Sprintf("%d.%s", inlineCount, picName), b, 0666)
					continue
				}
			}

			fmt.Printf("# %d%s : %s : %v\n", inlineCount, oExt, ct, ctpar)
			if bExtractBody {
				b, _ := ioutil.ReadAll(p.Body)
				os.WriteFile(fmt.Sprintf("%d%s", inlineCount, oExt), b, 0666)
			}
		case *eml.AttachmentHeader:
			act, actpar, _ := h.ContentType() // (t string, params map[string]string, err error)
			filename, _ := h.Filename()

			fmt.Printf("# %s : %s : %v\n", filename, act, actpar)
			if bExtractAttachMent {
				oFile, err := os.Create(filename)
				if err != nil {
					fmt.Println(err)
				}
				if _, err := io.Copy(oFile, p.Body); err != nil {
					fmt.Println(err)
				}
			}
		}
	}
}

// inline类型的附件，例如图片类型
func getInlineFilename(h *eml.InlineHeader) (string, error) {
	_, params, err := h.ContentDisposition()
	filename, ok := params["filename"]
	if !ok {
		_, params, err = h.ContentType()
		filename = params["name"]
	}
	return filename, err
}

// 将buf写入emlPath
func saveEml(buf *bytes.Buffer, emlPath string) int64 {
	f, err := os.OpenFile(emlPath, os.O_RDWR|os.O_CREATE, 0755)
	if err != nil {
		log.Fatal(err)
	}

	writeLen, _ := buf.WriteTo(f)

	if err := f.Close(); err != nil {
		log.Fatal(err)
	}
	return writeLen
}

func getFileSize(fileName string) int64 {
    fileInfo, err := os.Stat(fileName)
    if err!= nil {
        fmt.Println("获取文件信息时出错:", err)
        return 0
    }
    return fileInfo.Size()
}

func renameEml(emlName string) string {
	writeLen := getFileSize(emlName)

	r, _ := os.Open(emlName) // 读取eml
	mr, _ := eml.CreateReader(r)

	ei := NewEmlInfoFromMailHeader(mr.Header).SetLength(writeLen).SetUID(emlName)

	mr.Close()
	r.Close()

	lTime := ei.Date.Local()
	if err := os.Chtimes(emlName, lTime, lTime); err != nil {
		fmt.Println(err)
	}


	emlName1 := ei.GetFileName1()
	emlName2 := ei.GetFileName2()
	if emlName != emlName2 {
		fmt.Println(emlName, "->", emlName2)
		err := os.Rename(emlName, emlName1) // 第1次重命名
    	if err!= nil {
			fmt.Println(err)
		}
		err = os.Rename(emlName1, emlName2) // 第2次重命名
    	if err!= nil {
			fmt.Println(err)
		}
	}

	return ei.GetXMLLine()
}

func renameEmlFiles() {
    // 获取当前目录
    currentDir, err := os.Getwd()
    if err!= nil {
        log.Fatal(err)
    }

    // 遍历当前目录下的所有文件
    files, err := ioutil.ReadDir(currentDir)
    if err!= nil {
        log.Fatal(err)
    }

	var buf bytes.Buffer
    for _, file := range files {
        if filepath.Ext(file.Name()) == ".eml" {
			buf.WriteString(renameEml(file.Name()))
        }
    }

	err = os.WriteFile(XmlOfflinePATH, buf.Bytes(), os.ModePerm)
	if err != nil {
		log.Fatal(err)
	}
}

type EmlInfo struct {
	Date      time.Time
	Subject   string
	From      string
	To        string
	UID       string
	Length    int64
}

func (ei *EmlInfo) GetUnix() int64 {
	return ei.Date.Unix()
}

func (ei *EmlInfo) GetLocalTime() string {
	return ei.Date.Local().Format("2006-01-02_15.04.05")
}

func (ei *EmlInfo) GetNiceSubjct() string {
	return niceFileName(ei.Subject)
}

func (ei *EmlInfo) GetFileName1() string { // 一阶段文件名
	return fmt.Sprintf("%d_%s_%d.eml", ei.GetUnix(), ei.GetLocalTime(), ei.Length)
}

func (ei *EmlInfo) GetFileName2() string { // 二阶段文件名
	return fmt.Sprintf("%d_%s_%d_%s.eml", ei.GetUnix(), ei.GetLocalTime(), ei.Length, ei.GetNiceSubjct())
}

func (ei *EmlInfo) GetXMLLine() string { // 一行xml信息
	return fmt.Sprintf("<mail><ts>%d</ts><time>%s</time><size>%d</size><subject>%s</subject><from>%s</from><to>%s</to><mid>%s</mid></mail>\n", ei.GetUnix(), ei.GetLocalTime(), ei.Length, ei.Subject, ei.From, ei.To, ei.UID)
}

func (ei *EmlInfo) SetUID(iUID string) *EmlInfo {
	ei.UID = iUID
	return ei
}

func (ei *EmlInfo) SetLength(iLen int64) *EmlInfo {
	ei.Length = iLen
	return ei
}

func NewEmlInfo(iDate time.Time, iSubject string, iFrom string, iTo string, iUID string, iLen int64) *EmlInfo { // 创建EmlInfo
	return &EmlInfo{Date: iDate, Subject: iSubject, From: iFrom, To: iTo, UID: iUID, Length: iLen}
}

func NewEmlInfoFromMailHeader(h eml.Header) *EmlInfo { // 从header创建EmlInfo
	iSubject, _ := h.Text("subject")
	iDate, _    := emlDate(h.Get("date"))
	iFrom, _    := h.Text("from")
	iTo, _      := h.Text("to")
	return &EmlInfo{Date: iDate, Subject: iSubject, From: iFrom, To: iTo}
}

func NewEmlInfoFromMessageHeader(h msg.Header) *EmlInfo { // 从header创建EmlInfo
	iSubject, _ := h.Text("subject")
	iDate, _    := emlDate(h.Get("date"))
	iFrom, _    := h.Text("from")
	iTo, _      := h.Text("to")
	return &EmlInfo{Date: iDate, Subject: iSubject, From: iFrom, To: iTo}
}

