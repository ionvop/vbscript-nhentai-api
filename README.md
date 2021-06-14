# vbscript-nhentai-api
A VBScript-based API that retrieves the metadata of nhentai doujins. It doen't actually use any API gateways and the data are obtained from the source code of the doujin, so if there's any front-end changes to the site, the whole script might break. Ineffiecient, I know, nothing new. This was made for fun and wasn't made for practical use

This is still a work-in-progress. Right now it can only retrieve 16 different types of metadata (e.g., title, artists, tags) and 1 method `BookID(code)`

I'm still planning to add some updates and changes like instead of having `api.BookID(118282).cover` to retrieve the thumbnail image link while having `api.BookID(118282).Page(91).image` to retrieve image link of a specific page, it would only need `api.BookID(118282).image` to do so.
I also need to update the private function `funcListMidString()` because I actually found a more optimized version when I was working on the Javascript version of this



holy nuke code goes in, metadata goes out

<br>

## Methods

```
'Initialization
set api = new NhentaiApi

api.BookID(177013) 'Retrieve the metadata for this doujin

api.title
api.parodies
api.characters
api.tags
api.artists
api.groups
api.languages
api.categories
api.pages
api.uploaded
api.favorites
api.id
api.error
api.fullTitle
api.cover
api.statusCode

api.BookID(177013).title 'Returns "METAMORPHOSIS"
```
