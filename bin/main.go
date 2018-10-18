package main

import (
	"errors"
	"fmt"
	"html"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"time"

	outlooky "github.com/janmir/go-outlooky"
	"github.com/janmir/go-util"
	"github.com/janmir/pflag"
)

const (
	_post    = `☘`
	_nihongo = `[一-龯|぀-ゖ|ァ-ヺ|ｦ-ﾝ|ㇰ-ㇿ|Ａ-Ｚ|ａ-ｚ|０-９（）ー［］ー・]`

	_trys = 5
	_wait = 2000
)

var (
	flag       = pflag.CustomFlagSet("Outlooky", false, errors.New(""))
	temp       = flag.BoolP("temp", "x", false, "Temporary Testing Ground.")
	translator = flag.BoolP("translate", "t", true, "Translate Japanese Text to En.")
	unread     = flag.BoolP("unread", "u", true, "Changes will only apply to unread messages.")
	folder     = flag.StringP("folder", "f", "", "Comma-Serated paths to folder")

	nxp *regexp.Regexp
	pxp *regexp.Regexp

	procName   string
	removables = []string{
		"[External]",
	}
)

func init() {
	flag.Parse(os.Args[1:])

	var err error

	nxp, err = regexp.Compile(_nihongo)
	util.Catch(err)

	pxp, err = regexp.Compile(_post)
	util.Catch(err)

	procName = filepath.Base(os.Args[0])

	//Enable file logging
	util.EnableFileLogging()
}

func main() {
	//Check running instances
	if util.AmIRunning(procName) > 1 {
		util.Catch(errors.New("Instance is already running, exiting now... "))
	}

	switch {
	case *temp:
	case *translator:
		defer util.TimeTrack(time.Now(), "Translation")

		outlook := outlooky.Make()

		//Get all emails
		count, mails := getMails(outlook)
		util.Logger("Unread: ", count)

		if count == 0 {
			util.Catch(errors.New("Nothing to process"))
		}

		//Check for Japanese Text
		gt := NewTranslator()
		for _, v := range mails {
			subject := v.Subject
			og := v.Subject

			//Remove all removables
			for _, v := range removables {
				subject = strings.Replace(subject, v, "", -1)
			}

			//Check Subject
			if nxp.MatchString(subject) && !pxp.MatchString(subject) {
				// Translate
				util.Logger("Original: ", og)
				translated := gt.Translate(subject)
				util.Logger("Translated: ", translated)

				//un-escape translated string
				translated = html.UnescapeString(translated)

				//Clean whitespaces
				translated = strings.TrimSpace(translated)

				outlook.UpdateMail(v, outlooky.MailItem{
					Subject: fmt.Sprintf("⦗%s %s %s⦘ %s", _post, translated, _post, og),
				})
			} else {
				util.Logger("Skipped: ", subject)
			}
		}

	default:
		util.Red("Missing arguments, should either pass -t or -h/--help.")
	}
}

func getMails(outlook outlooky.Outlooky) (int, []outlooky.MailItem) {
	//pause first
	//util.Logger("Pausing...")
	//time.Sleep(time.Second * 10)

	try := _trys

	for try > 0 {
		count, mails := 0, []outlooky.MailItem{}

		if *folder == "" {
			count, mails = outlook.GetMails(outlooky.Inbox)
		} else {
			fs := strings.Split(*folder, ",")
			is := make([]interface{}, len(fs))
			for i, v := range fs {
				is[i] = v
			}

			count, mails = outlook.GetMails(is...)
		}

		if count == 0 {
			util.Catch(errors.New("No mails in inbox"))
		}

		//Filter unread
		if *unread {
			count, mails = outlook.ListMails(mails, true)

			if count == 0 {
				util.Logger("No unread mails")
			} else {
				return count, mails
			}
		}

		//Sleep
		try--
		util.Logger("Waiting...")
		time.Sleep(time.Millisecond * _wait)
	}
	return 0, []outlooky.MailItem{}
}
