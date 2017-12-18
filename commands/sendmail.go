package commands

import (
	"gopkg.in/urfave/cli.v1"
	"gopkg.in/urfave/cli.v1/altsrc"
	"../config"
	mailSendor "../sendMail"
)

func SendMail()  cli.Command {
	config := config.Config{}
	return cli.Command{
		Name:    "send",
		Aliases: []string{"s"},
		Usage:   "发送工资单邮件",
		Flags:[]cli.Flag{
			altsrc.NewIntFlag(cli.IntFlag{
				Name:        "excel.rowbegin",
				Value:       3,
				Usage:       "开始行号",
				Destination: &config.RowBegin,
			}),
			altsrc.NewStringFlag(cli.StringFlag{
				Name:        "excel.colbegin",
				Value:       "a",
				Usage:       "开始列号",
				Destination: &config.ColBegin,
			}),

			altsrc.NewStringFlag(cli.StringFlag{
				Name:        "excel.email",
				Value:       "a",
				Usage:       "邮箱所在列",
				Destination: &config.MailCol,
			}),

			altsrc.NewStringFlag(cli.StringFlag{
				Name:        "excel.name",
				Value:       "c",
				Usage:       "姓名所在列",
				Destination: &config.NameCol,
			}),

			altsrc.NewStringFlag(cli.StringFlag{
				Name:        "excel.sheet",
				Value:       "工资",
				Usage:       "Sheet名",
				Destination: &config.SheetName,
			}),
			altsrc.NewStringFlag(cli.StringFlag{
				Name:        "file",
				Value:       "",
				Usage:       "工资Excel文件",
				Destination: &config.ExcelPath,
			}),

			altsrc.NewStringFlag(cli.StringFlag{
				Name:        "mail.password",
				Value:       "",
				Usage:       "邮箱密码",
				Destination: &config.MailServer.Password,
			}),
			altsrc.NewStringFlag(cli.StringFlag{
				Name:        "mail.account",
				Value:       "",
				Usage:       "邮箱账号",
				Destination: &config.MailServer.Account,
			}),
			altsrc.NewStringFlag(cli.StringFlag{
				Name:        "mail.server",
				Value:      "smtp.exmail.qq.com",
				Usage:       "邮箱服务器",
				Destination: &config.MailServer.Server,
			}),
			altsrc.NewIntFlag(cli.IntFlag{
				Name:        "mail.port",
				Value:       25,
				Usage:       "邮箱服务器端口",
				Destination: &config.MailServer.Port,
			}),
		},

		Action: func(c *cli.Context) error {
			//fmt.Println(config)
			//fmt.Println(config.OutputMongo.MongoServer)
			sendMail := mailSendor.New(config)
			sendMail.Start()
			return nil
		},
	}
}
