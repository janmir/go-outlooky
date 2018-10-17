package outlooky

import (
	"testing"

	u "github.com/janmir/go-util"
)

func TestSingleGetMail(t *testing.T) {
	defer u.Recover()

	tree := outlook.GetMails("paulzu100@gmail.com")

	if len(tree.Leaves) > 0 {
		t.Fail()
	}

	tree = outlook.GetMails(_Inbox)

	if len(tree.Leaves) <= 0 {
		t.Fail()
	}

	tree = outlook.GetMails()

	if len(tree.Leaves) > 0 {
		t.Fail()
	}
}

func TestGetMail(t *testing.T) {
	defer u.Recover()

	tree := outlook.GetMails("paulzu100@gmail.com", "Custom")
	leaves := outlook.ListMails(tree, true)

	u.Logger(leaves)

	//make updates
	if len(leaves) > 0 {
		outlook.UpdateMail(leaves[0], MailItem{
			Subject:  "Hello",
			Body:     "Hello",
			HTMLBody: `<span style="font-size:11px;background-color:dimgray;color:white;font-family:Verdana;display:inline-block;padding:0px 5px;border-radius:5px;"> <img src="" alt="Hello there!"/> Hello </span>`,
		})
	} else {
		u.Logger("No Unread Messages.")
	}
}

func TestGetDefault(t *testing.T) {
	res := outlook.GetDefaultFolder(_Inbox)
	if res == nil {
		t.Fail()
	}
}

func TestGetCustom(t *testing.T) {
	folder := outlook.GetCustomFolder("paulzu100@gmail.com", "Custom", "Custom")
	if folder == nil {
		t.Fail()
	}

	// items := outlook.GetItems(folder)
	name := outlook.GetPropertyValue(folder, "Name").(string)
	u.Logger("Folder Name:", name)

	items := outlook.GetItems(folder)
	count := outlook.GetPropertyValue(items, "Count").(int32)
	u.Logger("Item Count:", count)

	if count <= 0 {
		t.Fail()
	}
}
func TestGetItems(t *testing.T) {
	folder := outlook.GetDefaultFolder(_Inbox)
	if folder == nil {
		t.Fail()
	}

	items := outlook.GetItems(folder)
	if items == nil {
		t.Fail()
	}

	count := outlook.GetPropertyValue(items, "Count").(int32)
	u.Logger("Item Count:", count)

	if count <= 0 {
		t.Fail()
	}
}

func TestGetPropertyValue(t *testing.T) {
	folder := outlook.GetDefaultFolder(_Inbox)
	if folder == nil {
		t.Fail()
	}

	items := outlook.GetItems(folder)
	if items == nil {
		t.Fail()
	}

	count := outlook.GetPropertyValue(items, "Count").(int32)
	if count <= 0 {
		t.Fail()
	}

	u.Logger("Item Count:", count)
}

func TestGetPropertyObject(t *testing.T) {
	folder := outlook.GetDefaultFolder(_Inbox)
	if folder == nil {
		t.Fail()
	}

	items := outlook.GetItems(folder)
	if items == nil {
		t.Fail()
	}

	item := outlook.GetPropertyObject(items, "Item", 1)
	if item == nil {
		t.Fail()
	}

	subject := outlook.GetPropertyValue(item, "Subject").(string)
	if subject == "" {
		t.Fail()
	}
	u.Logger("Mail Subject:", subject)
}

func TestSetPropertyValue(t *testing.T) {
	folder := outlook.GetDefaultFolder(_Inbox)
	if folder == nil {
		t.Fail()
	}

	items := outlook.GetItems(folder)
	if items == nil {
		t.Fail()
	}

	//sort first
	_, err := outlook.CallMethod(items, "Sort", "[ReceivedTime]", true)
	u.Catch(err)

	item := outlook.GetPropertyObject(items, "Item", 1)
	if item == nil {
		t.Fail()
	}

	subject := outlook.GetPropertyValue(item, "Subject").(string)
	if subject == "" {
		t.Fail()
	}
	u.Logger("Mail Subject:", subject)

	//update
	outlook.SetPropertyValue(item, "Subject", "New: "+subject)
	outlook.SaveItem(item)
}

func TestQuitApp(t *testing.T) {
	outlook.QuitApplication()
}
