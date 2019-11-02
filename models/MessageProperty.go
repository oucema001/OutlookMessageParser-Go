package models

//MessageProperty holds the type of data and the data
type MessageProperty struct {
	Class string
	Mapi  int64
	Data  interface{}
}
