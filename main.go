package main

import (
	"bufio"
	"crypto/md5"
	"encoding/base64"
	"encoding/hex"
	"flag"
	"fmt"
	"io"
	"log"
	"os"
	"regexp"
	"strconv"
	"strings"

	"github.com/emersion/go-message"
	"github.com/knadh/go-pop3"
	"golang.org/x/text/encoding/simplifiedchinese"
)

// WDir: 工作目录: eml保存目录
var (
	VerStr  string = "2023-07-07.11"
	WDir    string = "./emails/"
	LogPATH string = "00_bakEmail.md"
	LogStr  string = ""
	NowMode string = "list"
	NID     int    = 1
)

// NowMode: list, downAll, downOne, deleteOne

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
		fmt.Println("  ", os.Args[0], "-d 1                 下载序列号为1的邮件并释放附件")
		fmt.Println("  ", os.Args[0], "-e xx.eml            释放xx.eml中的附件")
		fmt.Println("  ", os.Args[0], "-rm 1                删除序列号为1的邮件")
		fmt.Println("  ", os.Args[0], "[-sub] -da 9         下载最近9封邮件")
		os.Exit(0)
	}
	bMakeDir2 := false // 以md5头两字符创建子文件夹
	flag.BoolVar(&bMakeDir2, "sub", bMakeDir2, "以md5头两字符创建子文件夹")
	bFixMD := false // 修复00_bakEmail.md，添加字段标题
	flag.BoolVar(&bFixMD, "fix", bFixMD, "修复00_bakEmail.md，添加字段标题")
	var downCount int
	flag.IntVar(&downCount, "da", -1, "倒序下载最近n封邮件")
	var listCount int
	flag.IntVar(&listCount, "l", -1, "倒序列出最近n封邮件的标题")
	var downIDX int
	flag.IntVar(&downIDX, "d", -1, "下载序列号为N的邮件并释放附件")
	var deleteIDX int
	flag.IntVar(&deleteIDX, "rm", -1, "删除序列号为N的邮件")
	emlPath := "xx.eml"
	flag.StringVar(&emlPath, "e", emlPath, "释放"+emlPath+"中的附件")
	flag.Parse() // 处理参数

	if "xx.eml" != emlPath { // 释放附件
		ExtractAttachmentsFromEml(emlPath)
		os.Exit(0)
	}

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
		envSV = "pop.163.com:995:1"
	}
	aEml := strings.Split(envSV, ":")
	bTLS := false
	if "" == aEml[0] || "" == aEml[1] || "" == aEml[2] {
		envSV = "pop.163.com:995:1"
		aEml = strings.Split(envSV, ":")
	}
	if "1" == aEml[2] {
		bTLS = true
	}
	nPort, err := strconv.Atoi(aEml[1])
	if err != nil {
		log.Fatal(err)
	}

	// 创建目录，进入目录
	var LogFile *os.File
	if "downAll" == NowMode || bFixMD {
		err = os.Mkdir(WDir, 0750)
		if err != nil && !os.IsExist(err) {
			log.Fatal(err)
		}
		os.Chdir(WDir)
		// 读取log，里面包含md5,len
		logBytes, _ := os.ReadFile(LogPATH)
		LogStr = string(logBytes)
		if bFixMD {
			os.Rename(LogPATH, LogPATH+".bak")
		}
		fLogtmp, err := os.OpenFile(LogPATH, os.O_RDWR|os.O_CREATE|os.O_APPEND, 0644)
		if err != nil {
			log.Fatal(err)
		}
		LogFile = fLogtmp
		// defer LogFile.Close()

		if bFixMD { // 读取*.eml，获取标题，添加到LogFile里面
			scanner := bufio.NewScanner(strings.NewReader(LogStr))
			for scanner.Scan() {
				line := scanner.Text()
				if strings.Contains(line, ",") {
					ff := strings.Split(line, ",")
					emlName := ff[0] + ".eml"
					if FileExist(emlName) {
						lineSub := strings.ReplaceAll(GetSubjectFromEml(emlName), ",", "，")
						fmt.Printf("- %s %s : %s\n", ff[0], ff[1], lineSub)
						if _, err := LogFile.WriteString(fmt.Sprintf("%s,%s,%s,\n", ff[0], ff[1], lineSub)); err != nil {
							LogFile.Close() // ignore error; Write error takes precedence
							log.Fatal(err)
						}
					}
				}
			}
			LogFile.Close()
			os.Exit(0)
		}
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

	//	Print the total number of messages and their size.
	count, size, _ := c.Stat()
	fmt.Println("# total messages =", count, ", size =", size)

	// Pull the list of all message IDs and their sizes.
	// msgs, _ := c.List(0)
	// for _, m := range msgs {
	// 	fmt.Println("- id =", m.ID, ", size =", m.Size)
	// }

	if "downOne" == NowMode { // downIDX
		if downIDX > count {
			downIDX = count
		}
		if 0 != count {
			m, _ := c.Retr(downIDX)
			nowSubject := emlSubjectDecode(m.Header.Get("subject"))
			fmt.Println("|", downIDX, ":", nowSubject)

			ExtractAttachmentFromEntity(m)
		}
	}
	if "deleteOne" == NowMode { // deleteIDX
		if deleteIDX > count {
			deleteIDX = count
		}
		if 0 != count {
			m, _ := c.Top(deleteIDX, 0) // 取头
			nowSubject := emlSubjectDecode(m.Header.Get("subject"))
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

		//		fmt.Println("- debug: count =", count, ", NID =", NID)
		// Pull all messages on the server. Message IDs go from 1 to N.
		// for id := NID; id <= count; id++ {
		for id := count; id >= NID; id-- {
			m, _ := c.Top(id, 0) // 取头
			nowSubject := emlSubjectDecode(m.Header.Get("subject"))
			fmt.Println("|", id, ":", nowSubject)

			if "downAll" == NowMode {
				buf, _ := c.RetrRaw(id)

				emlPath := fmt.Sprintf("%d.eml", id)
				f, err := os.OpenFile(emlPath, os.O_RDWR|os.O_CREATE, 0755)
				if err != nil {
					log.Fatal(err)
				}

				// m.WriteTo(f)
				writeLen, _ := buf.WriteTo(f)

				if err := f.Close(); err != nil {
					log.Fatal(err)
				}

				// 计算文件的md5，重命名
				nowMD5 := getFileMd5(emlPath)

				// 如果同文件存在，删除，否则，重命名
				if bHaveSameFile(nowMD5, fmt.Sprint(writeLen)) {
					os.Remove(emlPath)
					fmt.Printf("- %d : md5 = %s , len = %d\n", id, nowMD5, writeLen)
				} else { // 不存在，重命名
					dirName := ""
					if bMakeDir2 {
						dirName = nowMD5[0:2]
						err = os.Mkdir(dirName, 0750)
						if err != nil && !os.IsExist(err) {
							log.Fatal(err)
						}
						dirName = dirName + "/"
					}
					os.Rename(emlPath, dirName+nowMD5+".eml")
					fmt.Printf("+ %d : md5 = %s , len = %d\n", id, nowMD5, writeLen)
					if _, err := LogFile.WriteString(fmt.Sprintf("%s,%d,%s,\n", nowMD5, writeLen, nowSubject)); err != nil {
						LogFile.Close() // ignore error; Write error takes precedence
						log.Fatal(err)
					}
				}
			} // downAll
		}

		if "downAll" == NowMode {
			if err := LogFile.Close(); err != nil {
				log.Fatal(err)
			}
		}
	} // "downAll" || "list"

	fmt.Println("# total messages =", count, ", size =", size)
}

func GetSubjectFromEml(iPath string) string {
	// 打开eml文件
	fTest, err := os.OpenFile(iPath, os.O_RDONLY, 0644)
	if err != nil {
		fmt.Println("- Error @ GetSubjectFromEml : OpenFile :", err)
	}
	defer fTest.Close()

	// 解析eml
	m, err := message.Read(fTest)
	if err != nil {
		fmt.Println("- Error @ GetSubjectFromEml : message.Read :", err)
	}
	subRaw := m.Header.Get("subject")

	return emlSubjectDecode(subRaw)
}

func FileExist(filename string) bool {
	_, err := os.Stat(filename)
	return err == nil || os.IsExist(err)
}

func bHaveSameFile(iMD5 string, iLen string) bool {
	return strings.Contains(LogStr, iMD5+","+iLen+",")
}

func getFileMd5(filename string) string {
	pFile, err := os.Open(filename)
	if err != nil {
		fmt.Errorf("文件: %v 打开失败，err: %v", filename, err)
		return ""
	}
	defer pFile.Close()
	md5h := md5.New()
	io.Copy(md5h, pFile)
	return hex.EncodeToString(md5h.Sum(nil))
}

func GBK2UTF8(gbkStr string) string {
	utf8Str, _ := simplifiedchinese.GBK.NewDecoder().String(gbkStr)
	return utf8Str
}

func emlSubjectDecode(iSubject string) string {
	sSub := iSubject
	if strings.Contains(iSubject, "?B?") || strings.Contains(iSubject, "?b?") {
		sSub = ""
		matches := regexp.MustCompile("(?smi)=\\?([^\\?]+)\\?B\\?([^\\?]+)\\?=").FindAllStringSubmatch(iSubject, -1)
		for _, match := range matches {
			sDec, err := base64.StdEncoding.DecodeString(match[2])
			if err != nil {
				fmt.Println("- Error @ emlSubjectDecode : deCode Base64 :", iSubject)
			}
			encUP := strings.ToUpper(match[1])
			if encUP == "GBK" || encUP == "GB2312" || encUP == "GB18030" {
				sSub = sSub + GBK2UTF8(string(sDec))
			} else if encUP == "UTF-8" || encUP == "UTF8" {
				sSub = sSub + string(sDec)
			} else {
				// TODO
				fmt.Println("- Warning @ emlSubjectDecode : deCode Base64 :", iSubject)
				sSub = iSubject
				// sSub = sSub + string(sDec)
			}
		}
	} else if strings.Contains(iSubject, "?Q?") || strings.Contains(iSubject, "?q?") {
		sQStr := ""
		sSub = ""
		matches := regexp.MustCompile("(?smi)=\\?([^\\?]+)\\?Q\\?([^\\?]+)\\?=").FindAllStringSubmatch(iSubject, -1)
		for _, match := range matches {
			if strings.ToUpper(match[1]) != "UTF-8" {
				fmt.Println("- Error @ emlSubjectDecode : deCode Q 非UTF-8编码 :", match[1], sSub)
				continue
			}
			sSub = match[2]
			lenSub := len(sSub)

			// "_"替换为" ", "=FF=FF=FF":cnUTF8字符      "=":0x61  "_":0x95
			oStr := ""
			n := 0
			for {
				if n >= lenSub {
					break
				}
				if 61 == sSub[n] { // '='
					if n+3 > lenSub {
						fmt.Println("- Warning: 无法解码的神奇的字符串: ", sSub)
						break
					}
					/*
					   UTF-8:4字节
					   1: 0-127
					   2: 192-223, 128-191
					   3: 224-239, 128-191, 128-191
					   4: 240-247, 128-191, 128-191, 128-191
					*/
					// 测试首字符
					whiByte, _ := hex.DecodeString(string(sSub[n+1]) + string(sSub[n+2]))
					whiChar := whiByte[0]
					// fmt.Println("- n =", n, "lenSub =", lenSub, "whiChar =", whiChar)
					if whiChar < 128 { // 1字节 ascii
						enChar, _ := hex.DecodeString(string(sSub[n+1]) + string(sSub[n+2]))
						oStr = oStr + string(enChar)
						n = n + 3
					} else if whiChar < 224 { // 2字节
						thChar, _ := hex.DecodeString(string(sSub[n+1]) + string(sSub[n+2]) + string(sSub[n+4]) + string(sSub[n+5]))
						oStr = oStr + string(thChar)
						n = n + 6
					} else if whiChar < 240 { // 3字节: cn
						cnChar, _ := hex.DecodeString(string(sSub[n+1]) + string(sSub[n+2]) + string(sSub[n+4]) + string(sSub[n+5]) + string(sSub[n+7]) + string(sSub[n+8]))
						oStr = oStr + string(cnChar)
						n = n + 9
					} else {
						foChar, _ := hex.DecodeString(string(sSub[n+1]) + string(sSub[n+2]) + string(sSub[n+4]) + string(sSub[n+5]) + string(sSub[n+7]) + string(sSub[n+8]) + string(sSub[n+10]) + string(sSub[n+11]))
						oStr = oStr + string(foChar)
						n = n + 12
					}
				} else if 95 == sSub[n] { // '_'
					oStr = oStr + " "
					n = n + 1
				} else { // 未转义字符
					oStr = oStr + string(sSub[n])
					n = n + 1
				}
			}
			sQStr = sQStr + oStr
		}
		sSub = sQStr
	}
	return sSub
}

func ExtractAttachmentsFromEml(iPath string) {
	// 打开eml文件
	fTest, err := os.OpenFile(iPath, os.O_RDONLY, 0644)
	if err != nil {
		fmt.Println("- Error @ ExtractAttachmentsFromEml : OpenFile :", err)
	}
	defer fTest.Close()

	// 解析eml
	m, err := message.Read(fTest)
	if err != nil {
		fmt.Println("- Error @ ExtractAttachmentsFromEml : message.Read :", err)
	}
	subRaw := m.Header.Get("subject")
	// fmt.Println("- len(subject):", len(subRaw))
	// fmt.Println("- subject:", subRaw)

	fmt.Println("# subject:", emlSubjectDecode(subRaw))

	ExtractAttachmentFromEntity(m)
}

func ExtractAttachmentFromEntity(m *message.Entity) {
	mr := m.MultipartReader()
	for {
		nextEnt, err := mr.NextPart()
		if err != nil {
			if err == io.EOF {
				break
			}
			fmt.Println("- Error @ ExtractAttachmentFromEntity : NextPart :", err)
		}
		if "" != nextEnt.Header.Get("Content-Disposition") { // 附件
			_, aa, err := nextEnt.Header.ContentDisposition()
			if err != nil {
				fmt.Println("- Error @ ExtractAttachmentFromEntity : nextEnt.Header.ContentDisposition() :", err)
			}
			fileName := emlSubjectDecode(aa["filename"]) // 附件文件名
			// fmt.Println("- Content-Type:", nextEnt.Header.Get("Content-Type"))
			fmt.Println("- Content-Disposition:", nextEnt.Header.Get("Content-Disposition"))
			// fmt.Println("  - cs:", cs)
			// fmt.Println("  - aa:", aa)

			f, err := os.OpenFile(fileName, os.O_RDWR|os.O_CREATE, 0755)
			if err != nil {
				fmt.Println("- Error @ ExtractAttachmentFromEntity : OpenFile :", err)
			}
			fLen, err := io.Copy(f, nextEnt.Body)
			if err != nil {
				fmt.Println("- Error @ ExtractAttachmentFromEntity : io.Copy :", err)
			}
			fmt.Println("  - WriteTo :", fileName, ", FileLen :", fLen) // 附件文件名
		}
	}
}
