package sendMail

import (
	"../config"
	xlsx "github.com/360EntSecGroup-Skylar/excelize"
	"fmt"
	"strings"
	"os"
	"log"
	"io"
	"path"
	"path/filepath"
	"gopkg.in/gomail.v2"
)

type SendMail struct {
	config config.Config
}

func New(config config.Config) *SendMail {
	return &SendMail{
		config:config,
	}
}

func (this*SendMail) CopyTempfile(filename string)  {
	sFile, err := os.Open("./template.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	defer sFile.Close()

	eFile, err := os.Create(filename)
	if err != nil {
		log.Fatal(err)
	}
	defer eFile.Close()

	_, err = io.Copy(eFile, sFile) // first var shows number of bytes
	if err != nil {
		log.Fatal(err)
	}

	err = eFile.Sync()
	if err != nil {
		log.Fatal(err)
	}
}

func (this *SendMail)sendMail(attach string,subject string,to string,userName string)  {
	m := gomail.NewMessage()
	m.SetHeader("From", this.config.MailServer.Account)
	m.SetHeader("To", to)
	m.SetHeader("Subject", subject)
	m.SetBody("text/html", userName+":<br/>您好!<br/>"+subject+"工资单请查收!")

	_,fname := path.Split(attach)

	h := make(map[string][]string, 0)
	h["Content-Type"] = []string{`application/octet-stream; charset=utf-8; name="` + fname + `"`} //要设置这个，否则中文会乱码
	fileSetting := gomail.SetHeader(h)
	m.Attach(attach,fileSetting)

	d := gomail.NewDialer(this.config.MailServer.Server, this.config.MailServer.Port, this.config.MailServer.Account, this.config.MailServer.Password)
	if err := d.DialAndSend(m); err != nil {
		log.Fatalln(err)
		//panic(err)
		os.Exit(-1)
	}

}

func (this *SendMail) Start() {

	newpath := filepath.Join(".", "temp")
	os.MkdirAll(newpath, os.ModePerm)
	excelFileName := this.config.ExcelPath
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		log.Fatal(err)
		return
	}
	rows := xlFile.GetRows(this.config.SheetName)

	colBegin := strings.ToUpper(this.config.ColBegin)
	cb := int(colBegin[0])
	_,fname := path.Split(this.config.ExcelPath)
	finfo, err := os.Stat(this.config.ExcelPath)

	for i := this.config.RowBegin; i < len(rows); i++ {
		to := xlFile.GetCellValue(this.config.SheetName,this.config.MailCol+fmt.Sprintf("%d",i))
		userName := xlFile.GetCellValue(this.config.SheetName,this.config.NameCol+fmt.Sprintf("%d",i))
		filename := "./template.xlsx"
		dstFile, err := xlsx.OpenFile(filename)
		if err != nil {
			log.Fatal("xlsx.OpenFile")
			return
		}

		for j := cb; j <= int('A')+52;  j++{
			col :=  string(j)
			if j > int('Z') {
				col = "A"+ string(j-int('Z')+int('A')-1)
			}

			axis := col+fmt.Sprintf("%d",i)

			v1 := xlFile.GetCellValue(this.config.SheetName,col+"1");
			v2 := xlFile.GetCellValue(this.config.SheetName,col+"2");
			if v1 == "" && v2 == "" {
				break
			}
			//fmt.Printf("%s %s %s",col,v1,v2)

			value := xlFile.GetCellValue(this.config.SheetName,axis)
			dstFile.SetCellValue(this.config.SheetName,col+fmt.Sprintf("%d",this.config.RowBegin),value)
		}

		dstFile.SaveAs("temp/"+userName + fname)
		this.sendMail("temp/"+userName + fname,finfo.Name()[:len(finfo.Name())-5],to,userName)
	}

	os.RemoveAll("temp/")
	os.MkdirAll("temp/",os.ModePerm)

	log.Printf("处理成功!!!!")
}