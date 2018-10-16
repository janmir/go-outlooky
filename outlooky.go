package outlooky

import (
	"errors"
	"time"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	u "github.com/janmir/go-util"
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

/***************************************
	Outlooky-looky Functions
****************************************/

//GetMails ...
func (out Outlooky) GetMails(arg ...interface{}) Tree {
	defer u.TimeTrack(time.Now(), "GetMails")

	//get inbox
	var (
		inbox  *ole.IDispatch
		tree   = Tree{}
		length = len(arg)
		name   = ""
	)

	if length == 0 {
		return tree
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

	tree.Handle = inbox
	tree.Name = name
	tree.Leaves = out.GetLeaf(tree, MailItem{}, true)

	u.Logger("Fetched: ", len(tree.Leaves))

	return tree
}

//ListMails list/filter mail, read/unread
func (out Outlooky) ListMails(tree Tree, unread bool) []MailItem {
	defer u.TimeTrack(time.Now(), "ListMails")

	newList := make([]MailItem, 0)

	for _, v := range tree.Leaves {
		item := v.(MailItem)
		if item.UnRead == unread {
			newList = append(newList, item)
		}
	}

	return newList
}

//updateMail, subject/body

/***************************************
	Utility Functions
****************************************/

//GetLeaf returns a single item of type interface
func (out Outlooky) GetLeaf(tree Tree, identifier Leaf, sort bool) []interface{} {
	leaves := make([]interface{}, 0)

	branch := out.GetItems(tree.Handle)

	//sort
	if sort {
		_, err := out.CallMethod(branch, "Sort", "[ReceivedTime]", true)
		u.Catch(err)
	}

	//traverse
	count := int(out.GetPropertyValue(branch, "Count").(int32))

	//set limit
	count = u.Min(_Limit, count)

	//Traverse
	for i := 1; i <= count; i++ {
		leaf := out.GetItem(branch, i)
		leaves = append(leaves, identifier.Unmarshal(leaf))
	}

	return leaves
}

//get flags
//set flags

//GetDefaultFolder ...
func (out Outlooky) GetDefaultFolder(id int) *ole.IDispatch {
	folder, err := out.CallMethod(out.api, "GetDefaultFolder", id)
	u.Catch(err)

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
