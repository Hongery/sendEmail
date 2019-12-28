package routers

import (
	"sendEmail/controllers"
	"github.com/astaxie/beego"
)

func init() {
    beego.Router("/", &controllers.MainController{},"get:Get;post:Post")
	beego.Router("/success", &controllers.IndexController{},"get:Get")
}
