package models

type EmailParam struct {
	// ServerHost 邮箱服务器地址，如腾讯qq邮箱为smtp.qq.com
	ServerHost string
	// ServerPort 邮箱服务器端口，如腾讯qq邮箱为465
	ServerPort int
	// FromEmail　发件人邮箱地址
	FromEmail string
	// FromPasswd 发件人邮箱授权码（注意，这里是明文形式），TODO：如果设置成密文？
	FromPasswd string
	// Toers 接收者邮件，如有多个，则以英文逗号(“,”)隔开，不能为空
	Toers string
	// CCers 抄送者邮件，如有多个，则以英文逗号(“,”)隔开，可以为空
	CCers string
}