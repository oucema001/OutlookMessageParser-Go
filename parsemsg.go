package parse

import (
	"bytes"
	"encoding/binary"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/oucema001/OutlookMessageParser-go/models"
	"github.com/richardlehane/mscfb"
	"golang.org/x/net/html/charset"
)

const PropsKey = "__properties_version1.0"

// PropertyStreamPrefix is the prefix used for a property stream in the msg binary
const PropertyStreamPrefix = "__substg1.0_"

// ReplyToRegExp is a regex to extract the reply to header
const ReplyToRegExp = "^Reply-To:\\s*(?:<?(?<nameOrAddress>.*?)>?)?\\s*(?:<(?<address>.*?)>)?$"

//AnalyzeMsgFile analyzes the msg file and sets the properties
func AnalyzeMsgFile(file string) (res *models.Message, err error) {
	res = &models.Message{}
	f, err := os.Open(file)
	if err != nil {
		return nil, err
	}
	doc, err := mscfb.New(f)
	if err != nil {
		return nil, err
	}
	err = checkEntries(doc, res)
	//fmt.Println(res)
	if err != nil {
		return nil, err
	}
	return res, nil
}

func checkEntries(doc *mscfb.Reader, res *models.Message) error {
	var err error
	for entry, err := doc.Next(); err == nil; entry, err = doc.Next() {
		if strings.HasPrefix(entry.Name, PropertyStreamPrefix) {
			msg := outlookMessageProperty(entry)
			res.SetProperties(msg)
		}
	}
	return err
}

func outlookMessageProperty(entry *mscfb.File) models.MessageProperty {
	analysis := analyzeEntry(entry)
	data := getData(entry, analysis)
	messageProperty := models.MessageProperty{
		Class: analysis.Class,
		Mapi:  analysis.Mapi,
		Data:  data,
	}
	return messageProperty
}

func getData(entry *mscfb.File, info models.OutlookMessageInformation) interface{} {
	if info.Class == "" {
		return "class null"
	}
	mapi := info.Mapi
	switch mapi {
	case -1:
		return "-1"
	case 0x1e:
		bytes2 := make([]byte, entry.Size)
		entry.Read(bytes2)
		read, _ := charset.NewReader(bytes.NewReader(bytes2), "ISO-8859-1")
		if read != nil {
			resu, _ := ioutil.ReadAll(read)
			_ = resu

			return string(resu)
		}
		//return "type : : : 0X1E"
	case 0x1f:
		bytes2 := make([]byte, entry.Size)
		entry.Read(bytes2)
		//fmt.Println("++++++++++++++++++++", string(bytes2))
		//	read, _ := charset.NewReader(bytes.NewReader(bytes2), "utf-16")
		runes := make([]rune, len(bytes2)/2)
		c := 0
		for i := 0; i < len(bytes2)-1; i = i + 2 {
			ch := (int)(bytes2[i+1])
			cl := (int)(bytes2[i]) & 0xff
			runes[c] = (rune)((ch << 8) + cl)
			c++
		}
		return string(runes)
	case 0x102:
		return "type : : : 0x102"
	case 0x40:
		bytes := make([]byte, entry.Size)
		entry.Read(bytes)
		buf := make([]byte, 8, 8)
		if len(bytes) > 0 {
			buf = bytes[0 : len(bytes)-1]
			a := binary.BigEndian.Uint64(buf)
			a /= 10000
			a -= 11644473600000
			//return time.Date(a).String()
			return "type : : : 0x40" + time.Unix(0, int64(a)).String()
		}
		return "type : : : 0x40"
	default:
		return "default mapi : " + string(mapi)
	}
	return ""
}

//getEntriesFromDoc get properties from directory entry embedded TODO and TOCOMPLETE
func getEntriesFromDoc(entry *mscfb.File) []mscfb.File {
	result := make([]mscfb.File, 2)
	headerLength := 4
	flagsLength := 4
	bufHeader := make([]byte, headerLength)
	//bufHeader = buf[0:headerLength]
	var header strings.Builder
	//fmt.Println("hello")

	for k, err := entry.Read(bufHeader); err == nil && k > 0 && len(bufHeader) == headerLength; k, err = entry.Read(bufHeader) {
		//fmt.Println(err)
		//fmt.Println("hello")

		b, _ := entry.Seek(4, 1)

		_ = b
		header.Reset()
		//j := 0
		for i := len(bufHeader) - 1; i >= 0; i-- {
			header.WriteString(bytesToHex(bufHeader[i]))
			//fmt.Println("byte : ", bufHeader[j])
			//j++
		}

		class := header.String()[0:4]
		typeEntry := header.String()[4:]
		if typeEntry != "" {

			typeNumber, err := strconv.ParseInt(typeEntry, 16, 32)
			if err != nil {
				log.Print("Parse Error")
			}
			//fmt.Println("type number : ", typeNumber)
			if class != "0000" {
				bytes2 := make([]byte, flagsLength)

				entry.Read(bytes2)
				entry.Seek((int64)(len(bytes2)), 1)
				var k strings.Builder
				for _, b := range bytes2 {
					k.WriteString(bytesToHex(b))
				}
				//fmt.Println("flags :  ", k.String())
				if typeNumber == 0x48 || typeNumber == 0x1e || typeNumber == 0x1f || typeNumber == 0xd || typeNumber == 0x102 {
					//bytes1 := make([]byte, 4)
					//bytes2 = nil
					//fmt.Println("x")
					bytes2 = make([]byte, 4, 4)
					entry.Seek((int64)(cap(bytes2)), 1)
				} else if typeNumber == 0x3 || typeNumber == 0x4 || typeNumber == 0xa || typeNumber == 0xb || typeNumber == 0x2 {
					//SHORT
					// 4 bytes
					//bytes2 = nil
					//fmt.Println("y")
					bytes2 = make([]byte, 8, 8)
					entry.Read(bytes2[:4])
					entry.Seek((int64)(len(bytes2)), 1)
					entry.Read(bytes2[:8])
					entry.Seek((int64)(len(bytes2)), 1)
					//entry.Read(bytes2)
					//entry.Seek((int64)(len(bytes2)), 1)
				} else if typeNumber == 0x5 || typeNumber == 0x7 || typeNumber == 0x6 || typeNumber == 0x14 || typeNumber == 0x40 {
					// 8 bytes
					//bytes2 = nil
					//buf := make([]byte, 8, 8)
					bytes2 = make([]byte, 8, 8)
					entry.Read(bytes2)
					entry.Seek((int64)(cap(bytes2)), 1)
					//fmt.Println("8 bytes : ", len(bytes2))
					//fmt.Println("value : ", string(bytes2))
					//fmt.Println("number : ", typeNumber)
					//fmt.Println("buf size : ", len(buf))
					var SEC_TO_UNIX_EPOCH int64
					SEC_TO_UNIX_EPOCH = 11644473600
					var WINDOWS_TICK int64
					WINDOWS_TICK = 10000000
					a := int64(binary.BigEndian.Uint64(bytes2))

					fmt.Println("first : ", int64(a))
					a = (a / WINDOWS_TICK) - SEC_TO_UNIX_EPOCH
					b := a / 10000
					b = b - 11644473600000
					fmt.Println("milli : ", b)
					fmt.Println(time.Unix(a, 0).String())
				}

				if bytes2 != nil {

					name := "__substg1.0_" + header.String()
					var initial uint16
					initial = 95
					//var r io.Reade
					//a := smd.New()
					smd := new(mscfb.File)
					smd = &mscfb.File{
						Name:    name,
						Initial: initial,
						Path:    []string{header.String()},
						Size:    10000,
						//i:    0,
						//	i : 0,
						//r: &r,
					}
					//s,error:= mscfb.New()
					leng := (int64)(len("__substg1.0_"))

					_ = leng
					_ = smd
					if bytes2 != nil {
						fmt.Println(string(bytes2))
					}
				}
				bytes2 = nil
				bufHeader = nil
				bufHeader = make([]byte, headerLength)
				header.Reset()
			}

		}
	}
	return result
}

func analyzeEntry(entry *mscfb.File) models.OutlookMessageInformation {
	name := entry.Name
	res := models.OutlookMessageInformation{}
	if strings.HasPrefix(name, PropertyStreamPrefix) {
		var class string
		var typeEntry string
		var mapi int64
		val := name[len(PropertyStreamPrefix):]
		class = val[0:4]
		typeEntry = val[4:8]
		mapi, err := strconv.ParseInt(typeEntry, 16, 64)
		if err != nil {
			log.Println(err)
		}
		res = models.OutlookMessageInformation{
			Class: class,
			Mapi:  mapi,
		}
	}

	return res
}

func bytesToHex(bytes byte) string {
	//fmt.Printf("%02x", bytes&0xff)
	return fmt.Sprintf("%02x", bytes&0xff)
}
