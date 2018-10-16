package outlooky

import (
	"errors"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	u "github.com/janmir/go-util"
)

const (
	_Deleted  = 3  //Deleted items
	_Outbox   = 4  //Outbox
	_Sent     = 5  //Sent Items
	_Indox    = 6  //Inbox
	_Calendar = 9  //Calendar
	_Contacts = 10 //Contacts
	_Journal  = 11 //Journal
	_Notes    = 12 //Notes
	_Tasks    = 13 //Tasks

	_AppName = "Outlook.Application"
)

var (
	outlook Outlooky
)

//Outlooky ...
type Outlooky struct {
	handle *ole.IDispatch
	api    *ole.IDispatch
}

func init() {
	outlook = Outlooky{}
	ole.CoInitialize(0)

	unknown, err := oleutil.CreateObject(_AppName)
	u.Catch(err)

	handle, err := unknown.QueryInterface(ole.IID_IDispatch)
	u.Catch(err)

	api, err := oleutil.CallMethod(handle, "GetNamespace", "MAPI")
	u.Catch(err)

	//References
	outlook.handle = handle
	outlook.api = api.ToIDispatch()
}

/*
	Outlooky-looky
*/

//getmail, read/unread
//listmail, read/unread
//updateMail, subject/body

/*
	Utility Functions
*/

//per item
//get flags
//set flags

//per list
//sort

//GetDefaultFolder ...
func (out Outlooky) GetDefaultFolder(id int) *ole.IDispatch {
	folder, err := out.CallMethod(out.api, "GetDefaultFolder", id)
	u.Catch(err)

	return folder.ToIDispatch()
}

//GetCustomFolder ...
func (out Outlooky) GetCustomFolder(main string, subs ...string) *ole.IDispatch {
	folders := out.GetPropertyObject(out.api, "Folders") //Returns []Folder
	folder := out.GetItem(folders, main)

	for _, v := range subs {
		folders = out.GetPropertyObject(folder, "Folders")
		folder = out.GetItem(folders, v)
	}

	return folder //Returns Folder
}

//GetItem ...
// e.g Folder, Contact
func (out Outlooky) GetItem(folder *ole.IDispatch, arg ...interface{}) *ole.IDispatch {
	return out.GetPropertyObject(folder, "Item", arg...)
}

//GetItems ...
// e.g. _MailItem
func (out Outlooky) GetItems(folder *ole.IDispatch) *ole.IDispatch {
	items, err := out.CallMethod(folder, "Items")
	u.Catch(err)

	return items.ToIDispatch()
}

//GetPropertyValue ...
func (out Outlooky) GetPropertyValue(item *ole.IDispatch, name string, params ...interface{}) interface{} {
	prop, err := oleutil.GetProperty(item, name, params...)
	u.Catch(err)

	return prop.Value()
}

//GetPropertyObject ...
func (out Outlooky) GetPropertyObject(item *ole.IDispatch, name string, params ...interface{}) *ole.IDispatch {
	prop, err := oleutil.GetProperty(item, name, params...)
	u.Catch(err)

	if prop.VT != ole.VT_DISPATCH {
		u.Catch(errors.New("Not a dispatch object"))
	}

	return prop.ToIDispatch()
}

//SetPropertyValue ...
func (out Outlooky) SetPropertyValue(item *ole.IDispatch, name string, params ...interface{}) {
	_, err := oleutil.PutProperty(item, name, params...)
	u.Catch(err)

	//save changes
	out.SaveItem(item)
}

//SaveItem ...
func (out Outlooky) SaveItem(item *ole.IDispatch) {
	_, err := out.CallMethod(item, "Save")
	u.Catch(err)
}

//QuitApplication ...
func (out Outlooky) QuitApplication() {
	_, err := out.CallMethod(out.handle, "Quit")
	u.Catch(err)
}

//CallMethod ...
func (out Outlooky) CallMethod(handle *ole.IDispatch, method string, param ...interface{}) (*ole.VARIANT, error) {
	return oleutil.CallMethod(handle, method, param...)
}
