package models

import (
	"bufio"
	"fmt"
	"log"
	"net/textproto"
	"strconv"
	"strings"
	"time"
)

// Message is a struct that holds a structered result of parsing the entry
type Message struct {
	/**S
	 * The message class as defined in the .msg file.
	 */
	MessageClass string
	/**
	 * The message Id.
	 */
	MessageID string
	/**
	 * The address part of From: mail address.
	 */
	FromEmail string
	/**
	 * The name part of the From: mail address
	 */
	FromName string
	/**
	 * The address part of To: mail address.
	 */
	ToEmail string
	/**
	 * The name part of the To: mail address
	 */
	ToName string
	/**
	 * The address part of Reply-To header
	 */
	ReplyToEmail string
	/**
	 * The name part of Reply-To header
	 */
	ReplyToName string
	/**
	 * The S/MIME part of the S/MIME header
	 */
	//OutlookSmime smime;
	/**
	 * The mail's subject.
	 */
	Subject string
	/**
	 * The normalized body text.
	 */
	BodyText string
	/**
	 * The displayed To: field
	 */
	DisplayTo string
	/**
	 * The displayed Cc: field
	 */
	DisplayCc string
	/**
	 * The displayed Bcc: field
	 */
	DisplayBcc string

	/**
	 * The body in RTF format (if available)
	 */
	BodyRTF string

	/**
	 * The body in HTML format (if available)
	 */
	BodyHTML string

	/**
	 * The body in HTML format (converted from RTF)
	 */
	ConvertedBodyHTML string
	/**
	 * Email headers (if available)
	 */
	Headers string

	/**
	 * Email Date
	 */
	Date time.Time

	/**
	 * Client Submit Time
	 */
	ClientSubmitTime time.Time

	CreationDate time.Time

	LastModificationDate time.Time

	Properties map[int64]string
}

//SetProperties sets the message properties
func (res *Message) SetProperties(msgProps MessageProperty) {
	//res := Message{}
	name := msgProps.Class
	data := msgProps.Data
	if res.Properties == nil {
		res.Properties = make(map[int64]string, 2)
	}
	class, err := strconv.ParseInt(name, 16, 32)
	if err != nil {
		log.Print("Parse Error")
	}
	dataString := data.(string)
	switch class {
	case 0x1a:
		res.MessageClass = dataString
		break
	case 0x1035:
		res.MessageID = dataString
		break
	case 0x37:
		res.Subject = dataString
		break
	case 0xe1d:
		res.Subject = dataString
		break
	case 0xc1f: //SENDER EMAIL ADDRESS
	case 0x65: //SENT REPRESENTING EMAIL ADDRESS
	case 0x3ffa: //LAST MODIFIER NAME
	case 0x800d:
	case 0x8008:
		res.FromEmail = dataString
		break
	case 0x42:
		res.FromName = dataString
		break
	case 0x76:
		res.ToEmail = dataString
		break
	case 0x8000:
		res.ToEmail = dataString
		break
	case 0x3001:
		res.ToName = dataString
		break
	case 0xe04:
		res.DisplayTo = dataString
		break
	case 0xe03:
		res.DisplayCc = dataString
		break
	case 0xe02:
		res.DisplayBcc = dataString
		break
	case 0x1013:
		res.BodyHTML = dataString
		break
	case 0x1000:
		res.BodyText = dataString
		break
	case 0x1009:
		res.BodyRTF = dataString
		break
	case 0x7d:
		res.Headers = dataString
		break
	case 0x3007:
		//fmt.Println("hahhahahahah+++++++++++++++++++++++", dataString)
		//t, err := strconv.ParseInt(dataString, 2, 64)
		//if err != nil {
		//	log.Println(err)
		//}
		//res.CreationDate = time.Unix(0, t)
		res.CreationDate = getTimeFromString(string(dataString))

	case 0x3008:
		//fmt.Println("hahhahahahah+++++++++++++++++++++++", dataString)
		//t, err := strconv.ParseInt(dataString, 2, 64)
		//if err != nil {
		//	log.Println("custom", err)
		//}
		res.LastModificationDate = getTimeFromString(string(dataString))
	case 0x39:
		//fmt.Println("hahahhahahahha+++++++++++++++++++++++", dataString)
		//t, err := strconv.ParseInt(dataString, 2, 64)
		//if err != nil {
		//	log.Println("custom", err)
		//}
		//res.ClientSubmitTime = time.Unix(0, t)
		res.ClientSubmitTime = getTimeFromString(string(dataString))
	}
	res.Properties[class] = dataString
}

//GetHeaders returns headers
func (res *Message) GetHeaders() string {
	return res.Headers
}

//ParseHeaders returns a map of key value of headers
func (res *Message) ParseHeaders() map[string][]string {
	a := strings.NewReader(res.GetHeaders())
	tp := textproto.NewReader(bufio.NewReader(a))
	hdr, err := tp.ReadMIMEHeader()
	if err != nil {
		fmt.Println(err)
	}
	return hdr
}

func getTimeFromString(s string) (t time.Time) {
	if s == "" {
		return
	}
	var err error
	t, err = time.Parse(time.RFC1123Z, s)
	if err == nil {
		return t
	}

	t, err = time.Parse("Mon, 2 Jan 2006 15:04:05 -0700", s)
	return t
}
