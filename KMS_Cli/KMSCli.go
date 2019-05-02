package main

import (
	"bufio"
	"bytes"
	"encoding/xml"
	"flag"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"os/exec"
	"time"

	"github.com/axgle/mahonia"
)

type KMS_Servers struct {
	XMLName    xml.Name `xml:"KMS_Server_Site"`
	WinKey     string   `xml:"WinKey"`
	Office2016 string   `xml:"Office2016"`
	Office2019 string   `xml:"Office2019"`
	Servers    []server `xml:"Servers>Server"`
}

type server struct {
	XMLName     xml.Name `xml:"Server"`
	ID          string   `xml:"ID,attr"`
	Defalut     bool     `xml:"Defalut"`
	Name        string   `xml:"Name"`
	IP          string   `xml:"IP"`
	Description string   `xml:"Description"`
}

var logger *log.Logger
var CfgXML KMS_Servers

func main() {
	ConFile := flag.String("Config", "Config.XML", "指定XML配置文件。")
	flag.Parse()

	logfile, err := os.OpenFile("KMSCLI.log", os.O_APPEND|os.O_CREATE, 666)
	if err != nil {
		log.Fatalln("fail to create KMSCLI.log file!")
	}
	defer logfile.Close()
	logger = log.New(logfile, "", log.LstdFlags|log.Lshortfile) // 日志文件格式:log包含时间及文件行数
	logger.Println("==========  一次新的执行 ============")

	file, err := os.Open(*ConFile) // For read access.
	if err != nil {
		fmt.Printf("error: %v", err)
		return
	}
	defer file.Close()
	data, err := ioutil.ReadAll(file)
	if err != nil {
		fmt.Printf("error: %v", err)
		return
	}

	err = xml.Unmarshal(data, &CfgXML)
	if err != nil {
		fmt.Printf("error: %v", err)
		return
	}
	logger.Printf("获得配置文件KMS 站点设置 %v 个", len(CfgXML.Servers))
	//	fmt.Println(len(CfgXML.Servers))
	//	fmt.Println(CfgXML.Servers[1])
	mainmenu()
}

func mainmenu() {
	var Cmdout bytes.Buffer
	var Cmdstring string
	Site := 0
EXIT:
	for {
		cls := exec.Command("cmd", "/c", "cls")
		cls.Stdout = os.Stdout
		cls.Run()
		fmt.Printf("\x1b[1;32m%s\n", "=============================================================================")
		fmt.Printf("\x1b[1;31m%s\n", "              Windows10 与 Office 激活工具（Office2016、Office2019）")
		fmt.Printf("\x1b[1;32m%s\n", "=============================================================================")
		for i := 0; i < len(CfgXML.Servers); i++ {
			if CfgXML.Servers[i].Defalut {
				fmt.Print("    *")
				Site = i
			} else {
				fmt.Print("     ")
			}
			fmt.Printf("KMS 站点%d：%s \t[%s]\t%s\n", i+1, CfgXML.Servers[i].Name, CfgXML.Servers[i].IP, CfgXML.Servers[i].Description)
		}
		fmt.Printf("%s\n", "=============================================================================")
		fmt.Print("\t\t输入（1~9）选择KMS站点；\n")
		fmt.Printf("\t\t输入（a）设置WindowsKey %s\n", CfgXML.WinKey)
		fmt.Printf("\t\t输入（b）设置Office2016 %s\n", CfgXML.Office2016)
		fmt.Printf("\t\t输入（c）设置Office2019 %s\n", CfgXML.Office2019)
		fmt.Print("\t\t输入（w）激活或续期Windows\n")
		fmt.Print("\t\t输入（o）激活或续期Office \n")
		fmt.Print("\t\t输入（q）退出程序 \n")
		fmt.Printf("%s\n", "=============================================================================")
		fmt.Print("\t\t请输入选项：")
		input := bufio.NewScanner(os.Stdin)
		input.Scan()
		Sl := input.Bytes()[0]
		switch Sl {
		case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57: // 0~9
			Site = int(Sl) - 49
			logger.Printf("选择站点 %s  %d\n", string(Sl), Site)
		case 65, 97: // Aa
			logger.Printf("选择选项 %s\n", string(Sl))
			Cmdstring = fmt.Sprintf("cscript //Nologo C:\\Windows\\system32\\slmgr.vbs /ipk %s", CfgXML.WinKey)
			logger.Println(Cmdstring)
			setWKey := exec.Command("cmd", "/c", Cmdstring)
			setWKey.Stdout = &Cmdout
			setWKey.Run()
			logger.Print(ConvertToString(Cmdout.String(), "gbk", "utf-8"))
		case 66, 98: // Bb
			logger.Printf("选择选项 %s\n", string(Sl))
			Cmdstring = fmt.Sprintf("cscript //Nologo \"C:\\Program Files\\Microsoft Office\\Office16\\OSPP.VBS\" /inpkey:%s ", CfgXML.Office2016)
			logger.Println(Cmdstring)
			setOKey := exec.Command("cmd", "/c", Cmdstring)
			setOKey.Stdout = &Cmdout
			setOKey.Run()
			logger.Print(ConvertToString(Cmdout.String(), "gbk", "utf-8"))
		case 67, 99: // Cc
			logger.Printf("选择选项 %s\n", string(Sl))
			Cmdstring = fmt.Sprintf("cscript //Nologo \"C:\\Program Files\\Microsoft Office\\Office16\\OSPP.VBS\" /inpkey:%s ", CfgXML.Office2019)
			logger.Println(Cmdstring)
			setOKey := exec.Command("cmd", "/c", Cmdstring)
			setOKey.Stdout = &Cmdout
			setOKey.Run()
			logger.Print(ConvertToString(Cmdout.String(), "gbk", "utf-8"))
		case 87, 119: // Ww
			logger.Printf("选择选项 %s\n", string(Sl))
			// 设置KMS 站点
			Cmdstring = fmt.Sprintf("cscript //Nologo C:\\Windows\\system32\\slmgr.vbs /skms %s", CfgXML.Servers[Site].Name)
			logger.Println(Cmdstring)
			setWKey := exec.Command("cmd", "/c", Cmdstring)
			setWKey.Stdout = &Cmdout
			setWKey.Run()
			logger.Print(ConvertToString(Cmdout.String(), "gbk", "utf-8"))
			// 激活续期
			Cmdstring = "cscript //Nologo C:\\Windows\\system32\\slmgr.vbs /ato"
			logger.Println(Cmdstring)
			setWKey = exec.Command("cmd", "/c", Cmdstring)
			setWKey.Stdout = &Cmdout
			setWKey.Run()
			logger.Print(ConvertToString(Cmdout.String(), "gbk", "utf-8"))
			// 查看激活状态
			Cmdstring = "cscript //Nologo C:\\Windows\\system32\\slmgr.vbs /xpr"
			logger.Println(Cmdstring)
			setWKey = exec.Command("cmd", "/c", Cmdstring)
			setWKey.Stdout = &Cmdout
			setWKey.Run()
			logger.Print(ConvertToString(Cmdout.String(), "gbk", "utf-8"))

		case 79, 111: //Oo
			logger.Printf("选择选项 %s\n", string(Sl))
			Cmdstring = fmt.Sprintf("cscript //Nologo \"C:\\Program Files\\Microsoft Office\\Office16\\OSPP.VBS\" /sethst:%s ", CfgXML.Servers[Site].Name)
			logger.Println(Cmdstring)
			setOKey := exec.Command("cmd", "/c", Cmdstring)
			setOKey.Stdout = &Cmdout
			setOKey.Run()
			logger.Print(ConvertToString(Cmdout.String(), "gbk", "utf-8"))
			// 激活续期
			Cmdstring = "cscript //Nologo  \"C:\\Program Files\\Microsoft Office\\Office16\\OSPP.VBS\" /act"
			logger.Println(Cmdstring)
			setOKey = exec.Command("cmd", "/c", Cmdstring)
			setOKey.Stdout = &Cmdout
			setOKey.Run()
			logger.Print(ConvertToString(Cmdout.String(), "gbk", "utf-8"))
			// 查看激活状态
			Cmdstring = "cscript //Nologo  \"C:\\Program Files\\Microsoft Office\\Office16\\OSPP.VBS\" /dstatus"
			logger.Println(Cmdstring)
			setOKey = exec.Command("cmd", "/c", Cmdstring)
			setOKey.Stdout = &Cmdout
			setOKey.Run()
			logger.Print(ConvertToString(Cmdout.String(), "gbk", "utf-8"))
		case 81, 113: // Qq
			logger.Printf("选择选项 %s\n", string(Sl))
			break EXIT
		default:
			logger.Printf("选择 %s\n", string(Sl))
		}
		time.Sleep(3 * time.Second)
	}
	fmt.Print("\x1b[0m\n") // 设置显示恢复默认状态
	return
}

func ConvertToString(src string, srcCode string, tagCode string) string {
	srcCoder := mahonia.NewDecoder(srcCode)
	srcResult := srcCoder.ConvertString(src)
	tagCoder := mahonia.NewDecoder(tagCode)
	_, cdata, _ := tagCoder.Translate([]byte(srcResult), true)
	result := string(cdata)
	return result
}
