package outlooky

import (
	"errors"
	"fmt"
	"time"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/janmir/go-util"
)

const (
	_Limit    = 100 //1000
	_Deleted  = 3   //Deleted items
	_Outbox   = 4   //Outbox
	_Sent     = 5   //Sent Items
	_Inbox    = 6   //Inbox
	_Calendar = 9   //Calendar
	_Contacts = 10  //Contacts
	_Journal  = 11  //Journal
	_Notes    = 12  //Notes
	_Tasks    = 13  //Tasks

	_AppName = "Outlook.Application"

	//Inbox publicly available Inbox constant
	Inbox = _Inbox
)

var (
	outlook Outlooky
)

//Outlooky ...
type Outlooky struct {
	handle *ole.IDispatch
	api    *ole.IDispatch
}

//Tree Defines a tree
type Tree struct {
	Name   string
	Handle *ole.IDispatch
	Leaves []interface{}
}

//Leaf ...
type Leaf interface {
	Unmarshal(data *ole.IDispatch) interface{}
}

func init() {
	outlook = Outlooky{}
	ole.CoInitializeEx(0, 0)

	unknown, err := oleutil.CreateObject(_AppName)
	util.Catch(err)

	handle, err := unknown.QueryInterface(ole.IID_IDispatch)
	util.Catch(err)

	api, err := oleutil.CallMethod(handle, "GetNamespace", "MAPI")
	util.Catch(err)

	//References
	outlook.handle = handle
	outlook.api = api.ToIDispatch()
}

//Make Creates an instance of outlook
func Make() Outlooky {
	return outlook
}

/***************************************
	Outlooky-looky Functions
****************************************/

//GetMails ...
func (out Outlooky) GetMails(arg ...interface{}) (int, []MailItem) {
	defer util.TimeTrack(time.Now(), "GetMails")

	//get inbox
	var (
		inbox  *ole.IDispatch
		length = len(arg)
	)

	if length == 0 {
		return 0, []MailItem{}
	}

	if length == 1 {
		switch arg[0].(type) {
		case int:
			inbox = out.GetDefaultFolder(arg[0].(int))
		case string:
			inbox = out.GetCustomFolder(arg[0].(string))
		}
	} else {
		inbox = out.GetCustomFolder(arg[0].(string), arg[1:]...)
	}

	interfaces := out.GetLeaf(inbox, MailItem{}, true) //Returns []MailItem
	mails := make([]MailItem, len(interfaces))

	//Transfer
	for i, v := range interfaces {
		mails[i] = v.(MailItem)
	}

	util.Logger("GetMail Fetched: ", len(mails))

	return len(mails), mails
}

//ListMails list/filter mail, read/unread
func (out Outlooky) ListMails(mail []MailItem, unread bool) (int, []MailItem) {
	defer util.TimeTrack(time.Now(), "ListMails")

	newList := make([]MailItem, 0)

	for _, item := range mail {
		if item.UnRead == unread {
			newList = append(newList, item)
		}
	}

	return len(newList), newList
}

//UpdateMail ...
func (out Outlooky) UpdateMail(o, n MailItem) {
	defer util.TimeTrack(time.Now(), "UpdateMail")
	item := o.Data

	//Apply updates
	_ = oleutil.MustPutProperty(item, "Subject", n.Subject)
	// _ = oleutil.MustPutProperty(item, "Body", n.Body)
	// _ = oleutil.MustPutProperty(item, "HTMLBody", n.HTMLBody)

	//Save
	out.SaveItem(item)
}

/***************************************
	Utility Functions
****************************************/

//GetLeaf returns a single item of type interface
func (out Outlooky) GetLeaf(handle *ole.IDispatch, identifier Leaf, sort bool) []interface{} {
	leaves := make([]interface{}, 0)

	branch := out.GetItems(handle)

	//sort
	if sort {
		out.SortItems(branch, "[ReceivedTime]", true)
	}

	//traverse
	count := int(out.GetPropertyValue(branch, "Count").(int32))
	util.Logger("GetLeaf Count: ", count)

	if count > 0 {
		//set limit
		count = util.Min(_Limit, count)

		//Traverse
		for i := 1; i <= count; i++ {
			leaf := out.GetItem(branch, i)
			leaves = append(leaves, identifier.Unmarshal(leaf))
		}
	}

	return leaves
}

//get flags
//set flags

//GetDefaultFolder ...
func (out Outlooky) GetDefaultFolder(id int) *ole.IDispatch {
	folder, err := out.CallMethod(out.api, "GetDefaultFolder", id)
	util.Catch(err, "Get Default Folder Failed.")

	return folder.ToIDispatch()
}

//GetCustomFolder ...
func (out Outlooky) GetCustomFolder(main string, subs ...interface{}) *ole.IDispatch {
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
	util.Catch(err, "Failed retrieving items.")

	return items.ToIDispatch()
}

//GetPropertyValue ...
func (out Outlooky) GetPropertyValue(item *ole.IDispatch, name string, params ...interface{}) interface{} {
	prop, err := oleutil.GetProperty(item, name, params...)
	util.Catch(err, fmt.Sprintf(`Unable to get property value %s of value "%v"`, name, params))

	return prop.Value()
}

//GetPropertyObject ...
func (out Outlooky) GetPropertyObject(item *ole.IDispatch, name string, params ...interface{}) *ole.IDispatch {
	prop, err := oleutil.GetProperty(item, name, params...)
	util.Catch(err, fmt.Sprintf(`Unable to get property object %s of value "%v"`, name, params))

	if prop.VT != ole.VT_DISPATCH {
		util.Catch(errors.New("Not a dispatch object"))
	}

	return prop.ToIDispatch()
}

//SetPropertyValue ...
func (out Outlooky) SetPropertyValue(item *ole.IDispatch, name string, params ...interface{}) {
	_, err := oleutil.PutProperty(item, name, params...)
	util.Catch(err, fmt.Sprintf(`Unable to set property value %s to "%v"`, name, params))

	//save changes
	out.SaveItem(item)
}

//SortItems ...
func (out Outlooky) SortItems(items *ole.IDispatch, by string, desc bool) {
	_, err := out.CallMethod(items, "Sort", by, desc)
	util.Catch(err, "Sort failed execution")
}

//SaveItem ...
func (out Outlooky) SaveItem(item *ole.IDispatch) {
	_, err := out.CallMethod(item, "Save")
	util.Catch(err, "Unable to save the item.")
}

//QuitApplication ...
func (out Outlooky) QuitApplication() {
	_, err := out.CallMethod(out.handle, "Quit")
	util.Catch(err, "Unable to quit application.")
}

//CallMethod ...
func (out Outlooky) CallMethod(handle *ole.IDispatch, method string, param ...interface{}) (*ole.VARIANT, error) {
	return oleutil.CallMethod(handle, method, param...)
}
