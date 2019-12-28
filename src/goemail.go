package main

import (
    "github.com/go-gomail/gomail"
    "strings"
   // "time"
    "github.com/tealeg/xlsx"
    "log"
)
//goemail


type EmailParam struct {
    // ServerHost 邮箱服务器地址，如腾讯企业邮箱为smtp.exmail.qq.com
    ServerHost string
    // ServerPort 邮箱服务器端口，如腾讯企业邮箱为465
    ServerPort int
    // FromEmail　发件人邮箱地址
    FromEmail string
    // FromPasswd 发件人邮箱密码（注意，这里是明文形式），TODO：如果设置成密文？
    FromPasswd string
    // Toers 接收者邮件，如有多个，则以英文逗号(“,”)隔开，不能为空
    Toers string
    // CCers 抄送者邮件，如有多个，则以英文逗号(“,”)隔开，可以为空
    CCers string
}

// 全局变量，因为发件人账号、密码，需要在发送时才指定
// 注意，由于是小写，外面的包无法使用
/*var serverHost, fromEmail, fromPasswd string
var serverPort int*/

var m *gomail.Message

//初始化
func InitEmail(subject, body string,ep *EmailParam) {
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
    // 第三个参数为发件人别名，如"李大锤"，可以为空（此时则为邮箱名称）
    m.SetAddressHeader("From", ep.FromEmail, "CoderChain Information")
    m.SetHeader("Subject", subject)// 主题
    m.SetBody("text/html", body) // 正文
}

// SendEmail body支持html格式字符串
func SendEmail(ep *EmailParam) {
    /*m.SetHeader("Subject", subject)// 主题
    m.SetBody("text/html", body) // 正文*/
    d := gomail.NewPlainDialer(ep.ServerHost, ep.ServerPort, ep.FromEmail, ep.FromPasswd)
    // 发送
    /*if len(ep.Toers) != 0 {
        for _, tmp := range strings.Split(ep.Toers, ",") {
            m.SetHeader("To", strings.TrimSpace(tmp))
            *//*err := d.DialAndSend(m)
            if err != nil {
                panic(err)
            }*//*
        }
    }*/
    err := d.DialAndSend(m)
    if err != nil {
        panic(err)
    }
  /*  err := d.DialAndSend(m)
    if err != nil {
        panic(err)
    }*/
}

func main() {
    serverHost := "smtp.qq.com"//"smtp.exmail.qq.com"
    serverPort := 465
    fromEmail := "1**8@qq.com"//"test@latelee.org" 发送者邮箱
    fromPasswd := "hztebcspwpqfbafj" //"1qaz@WSX"  授权码
    myToers :="" //"li@latelee.org, latelee@163.com" // 逗号隔开 接收者邮箱
    myCCers := "huanggang1106@sina.com" //"readchy@163.com"抄送者
    subject := "这是主题"
    body := `这是正文<br>
            <h3>这是标题</h3>
             Hello <a href = "http://www.baidu.com">主页</a><br>`
    myEmail := &EmailParam {
        ServerHost: serverHost,
        ServerPort: serverPort,
        FromEmail:  fromEmail,
        FromPasswd: fromPasswd,
        Toers:      myToers,
        CCers:      myCCers,
    }

    TimeSettle(subject,body,myEmail)
}

func TimeSettle(subject,body string,myEmail *EmailParam) {
   /* d := time.Duration(time.Minute)
    t := time.NewTicker(d)
    defer t.Stop()
    for {
        currentTime := time.Now()
        if currentTime.Minute() == 21 {*/ // 8点发送   每个小时中的第17分钟

            //读取文件中的邮箱
            excelFileName := "C:/Users/huanggang/Desktop/fool.xlsx" //excel文件路径
            xlFile, err := xlsx.OpenFile(excelFileName)
            if err != nil {
                log.Panic(err)
            }
            for _, sheet := range xlFile.Sheets {
                for _, row := range sheet.Rows {
                    for _, cell := range row.Cells {
                        // myEmail.Toers :=cell.String()
                        myEmail.Toers=myEmail.Toers+cell.String()
                        /*
                        InitEmail(myEmail)
                        SendEmail(subject, body)*/

                    }
                }
            }

            InitEmail(subject, body,myEmail)
            SendEmail(myEmail)
            //time.Sleep(time.Minute)
        /*}
        <-t.C
    }*/

}

func main1111() {
    serverHost := "smtp.exmail.qq.com"
    serverPort := 465
    fromEmail := "test@latelee.org"
    fromPasswd := "1qaz@WSX"
    
    myToers := "li@latelee.org, latelee@163.com" // 逗号隔开
    myCCers := "" //"readchy@163.com"
    
    subject := "这是主题"
    body := `这是正文<br>
             Hello <a href = "http://www.latelee.org">主页</a>`
    // 结构体赋值
    myEmail := &EmailParam {
        ServerHost: serverHost,
        ServerPort: serverPort,
        FromEmail:  fromEmail,
        FromPasswd: fromPasswd,
        Toers:      myToers,
        CCers:      myCCers,
    }
    
    InitEmail(subject,body,myEmail)
    SendEmail(myEmail)
}
