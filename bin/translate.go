package main

import (
	"context"
	"fmt"
	"path/filepath"

	"cloud.google.com/go/translate"
	"github.com/janmir/go-util"
	"golang.org/x/text/language"
	"google.golang.org/api/option"
)

const (
	_authPath = "./auth.json"
)

//GTranslator ...
type GTranslator struct {
	client *translate.Client
	from   language.Tag
	to     language.Tag
	ctx    context.Context
}

//NewTranslator ...
func NewTranslator() GTranslator {
	// Sets the target language.
	target, err := language.Parse("en")
	util.Catch(err, fmt.Sprintf("Failed to parse target language: %v", err))

	source, err := language.Parse("ja")
	util.Catch(err, fmt.Sprintf("Failed to parse source language: %v", err))

	ctx := context.Background()

	path, err := util.GetCurrDir()
	util.Catch(err)

	file := filepath.Join(path, _authPath)
	// util.Logger("Auth Filepath: ", file)

	client, err := translate.NewClient(ctx, option.WithCredentialsFile(file))
	util.Catch(err, "Failed in initializing translation client.")

	return GTranslator{
		ctx:    ctx,
		from:   source,
		to:     target,
		client: client,
	}
}

//Translate returns translation object
func (g GTranslator) Translate(text string) []translate.Translation {
	util.Cyan("Translating: ", text)

	// Translates the text into English.
	translations, err := g.client.Translate(g.ctx, []string{text}, g.to, &translate.Options{
		Source: g.from,
	})
	util.Catch(err, fmt.Sprintf("Failed to translate text: %v", err))

	if len(translations) <= 0 {
		return []translate.Translation{}
	}

	util.Logger("Translations: %+v\n", translations)
	return translations
}

//TextTranslate returns the translated text
func (g GTranslator) TextTranslate(text string) string {
	// Translates the text into English.
	translations := g.Translate(text)
	return translations[0].Text
}
