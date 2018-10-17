package outlooky

import (
	"github.com/go-ole/go-ole"
)

//MailItem _MailItem
type MailItem struct {
	Data *ole.IDispatch

	//Actions                           //Returns an Actions collection that represents all the available actions for the item. Read-only.
	AlternateRecipientAllowed bool //Returns a Boolean (bool in C#) that is True if the mail message can be forwarded. Read/write.
	//Application                       //Returns an Application object that represents the parent Outlook application for the object. Read-only.
	//Attachments                       //Returns an Attachments object that represents all the attachments for the specified item. Read-only.
	AutoForwarded      bool   //A Boolean (bool in C#) value that returns True if the item was automatically forwarded. Read/write.
	AutoResolvedWinner bool   //Returns a Boolean (bool in C#) that determines if the item is a winner of an automatic conflict resolution. Read-only.
	BCC                string //Returns a String (string in C#) representing the display list of blind carbon copy (BCC) names for a MailItem. Read/write.
	BillingInformation string //Returns or sets a String (string in C#) representing the billing information associated with the Outlook item. Read/write.
	Body               string //Returns or sets a String (string in C#) representing the clear-text body of the Outlook item. Read/write.
	//BodyFormat                        //Returns or sets an OlBodyFormat constant indicating the format of the body text. Read/write.
	Categories string //Returns or sets a String (string in C#) representing the categories assigned to the Outlook item. Read/write.
	CC         string //Returns a String (string in C#) representing the display list of carbon copy (CC) names for a MailItem. Read/write.
	//Class                                    //Returns an OlObjectClass constant indicating the object's class. Read-only.
	Companies string //Returns or sets a String (string in C#) representing the names of the companies associated with the Outlook item. Read/write.
	//Conflicts                                //Return the Conflicts object that represents the items that are in conflict for any Outlook item object. Read-only.
	ConversationID    string //Returns a String (string in C#) that uniquely identifies a Conversation object that the MailItem object belongs to. Read-only.
	ConversationIndex string //Returns a String (string in C#) representing the index of the conversation thread of the Outlook item. Read-only.
	ConversationTopic string //Returns a String (string in C#) representing the topic of the conversation thread of the Outlook item. Read-only.
	//CreationTime                             //Returns a DateTime indicating the creation time for the Outlook item. Read-only.
	//DeferredDeliveryTime                     //Returns or sets a DateTime indicating the date and time the mail message is to be delivered. Read/write.
	DeleteAfterSubmit bool //Returns or sets a Boolean (bool in C#) value that is True if a copy of the mail message is not saved upon being sent, and False if a copy is saved. Read/write.
	//DownloadState                            //Returns a constant that belongs to the OlDownloadState enumeration indicating the download state of the item. Read-only.
	//EnableSharedAttachments                  //This object, member, or enumeration is deprecated and is not intended to be used in your code.
	EntryID string //Returns a String (string in C#) representing the unique Entry ID of the object. Read-only.
	//ExpiryTime                               //Returns or sets a DateTime indicating the date and time at which the item becomes invalid and can be deleted. Read/write.
	//FlagDueBy                                //This object, member, or enumeration is deprecated and is not intended to be used in your code.
	//FlagIcon                                 //This object, member, or enumeration is deprecated and is not intended to be used in your code.
	FlagRequest string //Returns or sets a String (string in C#) that indicates the requested action for a mail item. Read/write.
	//FlagStatus                               //This object, member, or enumeration is deprecated and is not intended to be used in your code.
	//FormDescription //Returns the FormDescription object that represents the form description for the specified Outlook item. Read-only.
	//GetInspector                             //Returns an Inspector object that represents an inspector initialized to contain the specified item. Read-only.
	//HasCoverSheet                            //This object, member, or enumeration is deprecated and is not intended to be used in your code.
	HTMLBody string //Returns or sets a String (string in C#) representing the HTML body of the specified item. Read/write.
	//Importance                               //Returns or sets an OlImportance constant indicating the relative importance level for the Outlook item. Read/write.
	//InternetCodepage                         //Returns or sets an Integer (int in C#) value that determines the Internet code page used by the item. Read/write.
	IsConflict bool //Returns a Boolean (bool in C#) that determines if the item is in conflict. Read-only.
	//IsIPFax                                  //This object, member, or enumeration is deprecated and is not intended to be used in your code.
	IsMarkedAsTask bool //Returns a Boolean (bool in C#) value that indicates whether the MailItem is marked as a task. Read-only.
	//ItemProperties                           //Returns an ItemProperties collection that represents all standard and user-defined properties associated with the Outlook item. Read-only.
	//LastModificationTime                     //Returns a DateTime specifying the date and time that the Outlook item was last modified. Read-only.
	//Links                                    //This object, member, or enumeration is deprecated and is not intended to be used in your code.
	//MAPIOBJECT                               //This object, member, or enumeration is deprecated and is not intended to be used in your code.
	//MarkForDownload                          //Returns or sets an OlRemoteStatus constant that determines the status of an item once it is received by a remote user. Read/write.
	MessageClass                      string //Returns or sets a String (string in C#) representing the message class for the Outlook item. Read/write.
	Mileage                           string //Returns or sets a String (string in C#) representing the mileage for an item. Read/write.
	NoAging                           bool   //Returns or sets a Boolean (bool in C#) value that is True to not age the Outlook item. Read/write.
	OriginatorDeliveryReportRequested bool   //Returns or sets a Boolean (bool in C#) value that determines whether the originator of the meeting item or mail message will receive a delivery report. Read/write.
	//OutlookInternalVersion                   //Returns an Integer (int in C#) value representing the build number of the Outlook application for an Outlook item. Read-only.
	OutlookVersion string //Returns a String (string in C#) indicating the major and minor version number of the Outlook application for an Outlook item. Read-only.
	//Parent                                   //Returns the parent Object of the specified object. Read-only.
	//Permission                               //Sets or returns an OlPermission constant that determines the permissions the recipients will have on the e-mail item. Read/write.
	//PermissionService                        //Sets or returns an OlPermissionService constant that determines the permission service that will be used when sending a message protected by Information Rights Management (IRM). Read/write.
	PermissionTemplateGUID string //Returns or sets a String (string in C#) value that represents the GUID of the template file to apply to the MailItem in order to specify Information Rights Management (IRM) permissions. Read/write.
	//PropertyAccessor                         //Returns a PropertyAccessor object that supports creating, getting, setting, and deleting properties of the parent MailItem object. Read-only.
	ReadReceiptRequested      bool   //Returns a Boolean (bool in C#) value that indicates True if a read receipt has been requested by the sender.
	ReceivedByEntryID         string //Returns a String (string in C#) representing the EntryID for the true recipient as set by the transport provider delivering the mail message. Read-only.
	ReceivedByName            string //Returns a String (string in C#) representing the display name of the true recipient for the mail message. Read-only.
	ReceivedOnBehalfOfEntryID string //Returns a String (string in C#) representing the EntryID of the user delegated to represent the recipient for the mail message. Read-only.
	ReceivedOnBehalfOfName    string //Returns a String (string in C#) representing the display name of the user delegated to represent the recipient for the mail message. Read-only.
	//ReceivedTime                             //Returns a DateTime indicating the date and time at which the item was received. Read-only.
	RecipientReassignmentProhibited bool //Returns a Boolean (bool in C#) that indicates True if the recipient cannot forward the mail message. Read/write.
	//Recipients                               //Returns a Recipients collection that represents all the recipients for the Outlook item. Read-only.
	ReminderOverrideDefault bool   //Returns or sets a Boolean (bool in C#) value that is True if the reminder overrides the default reminder behavior for the item. Read/write.
	ReminderPlaySound       bool   //Returns or sets a Boolean (bool in C#) value that is True if the reminder should play a sound when it occurs for this item. Read/write.
	ReminderSet             bool   //Returns or sets a Boolean (bool in C#) value that is True if a reminder has been set for this item. Read/write.
	ReminderSoundFile       string //Returns or sets a String (string in C#) indicating the path and file name of the sound file to play when the reminder occurs for the Outlook item. Read/write.
	//ReminderTime                             //Returns or sets a DateTime indicating the date and time at which the reminder should occur for the specified item. Read/write.
	//RemoteStatus                             //Returns or sets an OlRemoteStatus constant specifying the remote status of the mail message. Read/write.
	//ReplyRecipientNames                      //Returns a semicolon-delimited String (string in C#) list of reply recipients for the mail message. Read-only.
	//ReplyRecipients                          //Returns a Recipients collection that represents all the reply recipient objects for the Outlook item. Read-only.
	//RetentionExpirationDate                  //Returns a DateTime value that specifies the date when the MailItem object expires, after which the Messaging Records Management (MRM) Assistant will delete the item. Read-only.
	RetentionPolicyName string //Returns a String (string in C#) that specifies the name of the retention policy. Read-only.
	//RTFBody                                  //Returns or sets a byte array that represents the body of the Microsoft Outlook item in Rich Text Format. Read/write.
	Saved bool //Returns a Boolean (bool in C#) value that is True if the Outlook item has not been modified since the last save. Read-only.
	//SaveSentMessageFolder                    //Returns or sets a Folder object that represents the folder in which a copy of the e-mail message will be saved after being sent. Read/write.
	//Sender                                   //Returns or sets an AddressEntry object that corresponds to the user of the account from which the MailItem is sent. Read/write.
	SenderEmailAddress string //Returns a String (string in C#) that represents the e-mail address of the sender of the Outlook item. Read-only.
	SenderEmailType    string //Returns a String (string in C#) that represents the type of entry for the e-mail address of the sender of the Outlook item, such as 'SMTP' for Internet address, 'EX' for a Microsoft Exchange server address, etc. Read-only.
	SenderName         string //Returns a String (string in C#) indicating the display name of the sender for the Outlook item. Read-only.
	//SendUsingAccount                         //Returns or sets an Account object that represents the account under which the MailItem is to be sent. Read/write.
	//Sensitivity                              //Returns or sets a constant in the OlSensitivity enumeration indicating the sensitivity for the Outlook item. Read/write.
	Sent bool //Returns a Boolean (bool in C#) value that indicates if a message has been sent. Read-only.
	//SentOn                                   //Returns a DateTime indicating the date and time on which the Outlook item was sent. Read-only.
	SentOnBehalfOfName string //Returns a String (string in C#) indicating the display name for the intended sender of the mail message. Read/write.
	//Session                                  //Returns the NameSpace object for the current session. Read-only.
	Size      int32  //Returns an Integer (int in C#) value indicating the size (in bytes) of the Outlook item. Read-only.
	Subject   string //Returns or sets a String (string in C#) indicating the subject for the Outlook item. Read/write.
	Submitted bool   //Returns a Boolean (bool in C#) value that is True if the item has been submitted. Read-only.
	//TaskCompletedDate                        //Returns or sets a DateTime value that represents the completion date of the task for this MailItem. Read/write.
	//TaskDueDate                              //Returns or sets a DateTime value that represents the due date of the task for this MailItem. Read/write.
	//TaskStartDate                            //Returns or sets a DateTime value that represents the start date of the task for this MailItem object. Read/write.
	TaskSubject string //Returns or sets a String (string in C#) value that represents the subject of the task for the MailItem object. Read/write.
	To          string //Returns or sets a semicolon-delimited String (string in C#) list of display names for the To recipients for the Outlook item. Read/write.
	//ToDoTaskOrdinal                          //Returns or sets a DateTime value that represents the ordinal value of the task for the MailItem. Read/write.
	UnRead bool //Returns or sets a Boolean (bool in C#) value that is True if the Outlook item has not been opened (read). Read/write.
	//UserProperties                           //Returns the UserProperties collection that represents all the user properties for the Outlook item. Read-only.
	VotingOptions  string //Returns or sets a String (string in C#) specifying a delimited string containing the voting options for the mail message. Read/write.
	VotingResponse string //Returns or sets a String (string in C#) specifying the voting response for the mail message. Read/write.
}

//Unmarshal ...
func (m MailItem) Unmarshal(data *ole.IDispatch) interface{} {
	mm := MailItem{Data: data}
	mm.Subject = outlook.GetPropertyValue(data, "Subject").(string)
	mm.UnRead = outlook.GetPropertyValue(data, "UnRead").(bool)
	// mm.HTMLBody = outlook.GetPropertyValue(data, "HTMLBody").(string)
	// mm.Body = outlook.GetPropertyValue(data, "Body").(string)

	return mm
}
