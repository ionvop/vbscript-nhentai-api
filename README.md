# vbscript-nhentai-api
A VBScript-based makeshift API for nhentai using only their html data

holy nuke code goes in, details go out

<br>

## Classes

### cNhentaiApi

The first version of the class using let and get properties.

```
set api = new cNhentaiApi
api.BookID = 177013
wscript.echo api.getArtists 'Echoes "shindol"

api.BookID = 354125
wscript.echo api.getTitle 'Echoes "AzuLan Anime Ero Mousou Hon"
```

<br>

### cNhentaiApiv2

The second version of the class using functions and variables.\
Can be set similarly to the first version or can utilize method chaining.

```
set api2 = new cNhentaiApiv2
api2.BookID(177013)
wscript.echo api2.title 'Echoes "METAMORPHOSIS"

wscript.echo api2.(177013).artists 'Echoes "shindol"

api2.BookID(354125)
wscript.echo api2.artists 'Echoes ["noukatu", "minase kuru"]

wscript.echo api2.(354125).artists 'Echoes "AzuLan Anime Ero Mousou Hon"
