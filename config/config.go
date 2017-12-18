package config


type Config struct {

	ColBegin string
	RowBegin int
	MailCol string
	NameCol string
	SheetName string
	ExcelPath string
	MailServer MailServer
}


type MailServer struct {
	Server string
	Port int
	Account string
	Password string

}
