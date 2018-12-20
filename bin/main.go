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

	"github.com/dgraph-io/badger"
	outlooky "github.com/janmir/go-outlooky"
	"github.com/janmir/go-util"
	"github.com/janmir/pflag"
)

const (
	_marker  = `✱`
	_nihongo = `[一-龯|぀-ゖ|ァ-ヺ|ｦ-ﾝ|ㇰ-ㇿ|Ａ-Ｚ|ａ-ｚ|０-９（）ー［］ー・]`

	_trys = 5
	_wait = 2000

	_cacheStoreDir = "."
)

var (
	flag       = pflag.CustomFlagSet("Outlooky", false, errors.New(""))
	temp       = flag.BoolP("temp", "x", false, "Temporary Testing Ground.")
	translator = flag.BoolP("translate", "t", true, "Translate Japanese Text to English.")
	unread     = flag.BoolP("unread", "u", true, "Changes will only apply to unread messages.")
	folder     = flag.StringP("folder", "f", "", "Comma-Serated paths to folder")

	nxp *regexp.Regexp
	pxp *regexp.Regexp

	procName   string
	removables = []string{
		"[External]",
	}

	//database
	db *badger.DB
)

func init() {
	flag.Parse(os.Args[1:])

	var err error

	nxp, err = regexp.Compile(_nihongo)
	util.Catch(err)

	pxp, err = regexp.Compile(_marker)
	util.Catch(err)

	procName = filepath.Base(os.Args[0])

	//initialize the data store
	dir := util.Localize(_cacheStoreDir)

	opts := badger.DefaultOptions
	opts.Dir = dir
	opts.ValueDir = dir

	db, err = badger.Open(opts)
	util.Catch(err)
}

func main() {
	defer db.Close()

	if _debug {
		//Dev mode
	} else {
		//Release Mode
		util.EnableFileLogging()
	}

	//Check running instances
	if util.AmIRunning(procName) > 1 {
		db.Close()
		util.Catch(errors.New("Instance is already running, exiting now... "))
	}

	switch {
	case *temp:
		//do nothing, this is only used to test init
		//function contents
	case *translator:
		defer util.TimeTrack(time.Now(), "Translation")

		outlook := outlooky.Make()

		//Get all emails
		count, mails := getMails(outlook)
		util.Logger("Unread: ", count)

		if count == 0 {
			db.Close()
			util.Catch(errors.New("Nothing to process"))
		}

		//Check for Japanese Text
		gtranslator := NewTranslator()
		for _, v := range mails {
			subject := v.Subject
			og := v.Subject

			//Remove all removables
			for _, v := range removables {
				subject = strings.Replace(subject, v, "", -1)
			}

			//Check Subject if already translated or if it contains
			//japanese characters
			if nxp.MatchString(subject) && !pxp.MatchString(subject) {
				translated := getCache(og)
				//check from cache first
				if len(translated) > 0 {
					util.Logger("✱ Cache Translate ✱")
				} else {
					util.Logger("✱ Google Translate ✱")

					//if not in cache Translate using google translator
					util.Logger("Original ➜ ", og)
					translated = gtranslator.TextTranslate(subject)
					util.Logger("Translated ➜ ", translated)

					//un-escape translated string
					translated = html.UnescapeString(translated)

					//Clean pre/post whitespaces
					translated = strings.TrimSpace(translated)

					//add to cache
					err := db.Update(func(txn *badger.Txn) error {
						err := txn.Set([]byte(og), []byte(translated))
						return err
					})
					util.Catch(err)
				}

				outlook.UpdateMail(v, outlooky.MailItem{
					Subject: fmt.Sprintf("⦗%s %s %s⦘ %s", _marker, translated, _marker, og),
				})
			} else {
				util.Logger("Skipped ➜ ", subject)
			}
		}

	default:
		util.Red("Missing arguments, should either pass -t/--translate or -h/--help.")
	}
}

func getCache(key string) string {
	var valCopy []byte
	err := db.View(func(txn *badger.Txn) error {
		item, err := txn.Get([]byte(key))
		if err != nil {
			return err
		}

		err = item.Value(func(val []byte) error {
			// This func with val would only be called if item.Value encounters no error.

			// Copying or parsing val is valid.
			valCopy = append([]byte{}, val...)

			return nil
		})
		return err
	})

	if err != nil {
		switch err {
		case badger.ErrKeyNotFound:
			//key does not exist
			util.Logger("Error: Key %q not found.", key)
		}
	}

	return string(valCopy)
}

func getMails(outlook outlooky.Outlooky) (int, []outlooky.MailItem) {
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

		//decrement remaining trials
		try--

		//Sleep
		if try > 0 {
			util.Logger("Waiting...")
			time.Sleep(time.Millisecond * _wait)
		}
	}
	return 0, []outlooky.MailItem{}
}
