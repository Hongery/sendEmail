package controllers

import (
	"github.com/astaxie/beego"
	"github.com/go-gomail/gomail"
	"github.com/tealeg/xlsx"
	"log"
	"time"
	"sendEmail/models"
	"strings"
	"path"
	//"fmt"
	"fmt"
)
type MainController struct {
	beego.Controller
}

func (c *MainController) Get() {
	c.TplName = "index.html"
}

// 全局变量，因为发件人账号、密码，需要在发送时才指定
// 注意，由于是小写，外面的包无法使用
var m *gomail.Message

//初始化
func (c *MainController)InitEmail(subject,body string,ep *models.EmailParam) {
	ccers :=[]string{}
	m = gomail.NewMessage()
	//抄送列表
	if len(ep.CCers) != 0 {
		for _, tmp := range strings.Split(ep.CCers, ",") {
			ccers = append(ccers, strings.TrimSpace(tmp))
		}
		m.SetHeader("Cc", ccers...)
	}
	// 发件人
	// 第三个参数为发件人别名，如"CoderChain"，可以为空（此时则为邮箱名称）
	m.SetAddressHeader("From", ep.FromEmail, "CoderChain Information")
	m.SetHeader("Subject", subject) // 主题
	m.SetBody("text/html", body)// 正文 body支持html格式字符串
}

// SendEmail
func (c *MainController)SendEmail(ep *models.EmailParam) {
	d := gomail.NewPlainDialer(ep.ServerHost, ep.ServerPort, ep.FromEmail, ep.FromPasswd)
	// 发送
	if len(ep.Toers) != 0 {
		for _, tmp := range strings.Split(ep.Toers, ",") {
			m.SetHeader("To", strings.TrimSpace(tmp))
			err := d.DialAndSend(m)
			if err != nil {
				s :=fmt.Sprint(err)
				c.Ctx.WriteString(s)
				//fmt.Println(err)
				//panic(err)
			}
		}
	}else{
		return
	}
}

func (c *MainController) Post(){
	serverHost := c.GetString("serverHost") //"smtp.qq.com"
	serverPort,_ := c.GetInt("serverPort") //465 返回int err
	fromEmail :=  c.GetString("fromEmail") //"1281185088@qq.com"// 发送者邮箱
	fromPasswd := c.GetString("fromPasswd")//"hztebcspwpqfbafj"// 授权码
	myToers :=""// 逗号隔开 接收者邮箱
	myCCers := c.GetString("myCCers")//"hongery@yeah.net" //抄送者邮箱
	//1.拿到要发送的主题和内容
	subject := c.GetString("subject") //"这是主题"
	body := c.GetString("content")
	if serverHost == " " || serverPort == 0 || fromPasswd == "" || fromEmail== "" || subject == "" || body == ""  {
		c.Ctx.WriteString("除抄送外不能有空")
		//c.Redirect("/",302)
	}else {
		//文件上传功能
		f, h, err := c.GetFile("uploadname")
		//defer f.Close()
		//1.限定格式 .xlsx
		fileext := path.Ext(h.Filename) //取出后缀
		beego.Info(fileext)
		if fileext != ".xlsx" {
			beego.Info("上传文件格式错误")
			return
		}
		//2.限制大小
		/*if h.Size > 1000000000{
		beego.Info("上传文件过大")
		return
	}*/
		//3.对文件重新命名，防止重复
		filename := time.Now().Format("2006-01-02") + fileext //6-1-2 3:4:5
		if err != nil {
			beego.Info("上传文件失败")
			/*fmt.Println("getfile err",err)*/
		} else {
			//保存文件
			c.SaveToFile("uploadname", "./static/"+filename)
		}

		myEmail := &models.EmailParam{
			ServerHost: serverHost,
			ServerPort: serverPort,
			FromEmail:  fromEmail,
			FromPasswd: fromPasswd,
			Toers:      myToers,
			CCers:      myCCers,
		}
		//发送邮件
		c.TimeSettle(filename, subject, body, myEmail)
		beego.Info("success")
		//发送成功
		c.Redirect("/success", 302)
		defer	f.Close()
	}

}
func (c *MainController)TimeSettle(filename,subject,body string,myEmail *models.EmailParam) {
			//读取文件中的邮箱
			excelFileName := "./static/"+filename //"C:/Users/huanggang/Desktop/fool.xlsx" //excel文件路径
			//fmt.Println(excelFileName)
			xlFile, err := xlsx.OpenFile(excelFileName)
			if err != nil {
				log.Panic(err)
			}
			//读取xlsx文件
			for _, sheet := range xlFile.Sheets {
				for _, row := range sheet.Rows {
					for _, cell := range row.Cells {
						myEmail.Toers=myEmail.Toers+cell.String()
					}
				}
			}
			c.InitEmail(subject, body,myEmail)
			c.SendEmail(myEmail)
}