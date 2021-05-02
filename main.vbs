set objShell = CreateObject("wscript.shell")
set objFile = CreateObject("Scripting.FileSystemObject")
set objHTML = CreateObject("htmlfile")
set objHTTP = CreateObject("MSXML2.XMLHTTP")

strDirectory = objShell.CurrentDirectory

function funcMidString(fInput, fLeft, fRight)
    if instr(fInput, fLeft) or instr(fInput, fRight) then
    else
        funcMidString = false
        exit function
    end if

    funcMidString = mid(fInput, instr(fInput, fLeft) + len(fLeft))

    if instr(funcMidString, fRight) then
    else
        funcMidString = false
        exit function
    end if

    funcMidString = left(funcMidString, instr(funcMidString, fRight) - 1)
end function

function funcListMidString(fInput, fLeft, fRight)
    dim varPosition, strItem, strList
    varPosition = 0
    strItem = ""
    strList = ""

    do
        if instr(varPosition + 1, fInput, fLeft) then
        else
            exit do
        end if

        varPosition = instr(varPosition + 1, fInput, fLeft)
        fInput = mid(fInput, varPosition)
        varPosition = 1

        if instr(fInput, fRight) then
        else
            exit do
        end if

        strItem = funcMidString(fInput, fLeft, fRight)

        if strItem = false then
            exit do
        end if

        strList = strList & strItem & vbCrlf
    loop

    if len(strList) > 2 then
        funcListMidString = left(strList, len(strList) - 2)
    end if
end function

function funcNhentaiGetInfo(fCode)
    dim strTitle, strParodies, strCharacters, strTags, strArtists, strGroups, strLanguages, strCategories, varPages, strUploaded, varFavorites, strFullTitle, strCoverImage
    objHTTP.open "GET", "https://nhentai.net/g/" & fCode, false
    objHTTP.send
    strData = objHTTP.responsetext
    objFile.createtextfile(strDirectory & "/" & fCode & ".html", true, true).writeline(strData)

    if instr(strData, "Looks like what you're looking for isn't here.") then
        funcNhentaiGetInfo = 404
        exit function
    end if

    strTitle = funcMidString(strData, "<meta itemprop=""name"" content=""", """ />")
    strParodies = funcMidString(strData, "Parodies:", "Characters:")
    strParodies = funcListMidString(strParodies, "<span class=""name"">", "</span>")
    strCharacters = funcMidString(strData, "Characters:", "Tags:")
    strCharacters = funcListMidString(strCharacters, "<span class=""name"">", "</span>")
    'strTags = funcMidString(strData, "Tags:", "Artists:")
    'strTags = funcListMidString(strTags, "<span class=""name"">", "</span>")
    strTags = funcMidString(strData, "<meta name=""twitter:description"" content=""", """ />")
    strTags = replace(strTags, ", ", vbCrlf)
    strArtists = funcMidString(strData, "Artists:", "Groups:")
    strArtists = funcListMidString(strArtists, "<span class=""name"">", "</span>")
    strGroups = funcMidString(strData, "Groups:", "Languages:")
    strGroups = funcListMidString(strGroups, "<span class=""name"">", "</span>")
    strLanguages = funcMidString(strData, "Languages:", "Categories:")
    strLanguages = funcListMidString(strLanguages, "<span class=""name"">", "</span>")
    strCategories = funcMidString(strData, "Categories:", "Pages:")
    strCategories = funcListMidString(strCategories, "<span class=""name"">", "</span>")
    varPages = funcMidString(strData, "Pages:", "Uploaded:")
    varPages = funcListMidString(varPages, "<span class=""name"">", "</span>")
    strUploaded = funcMidString(strData, "Uploaded:", "ript>")
    strUploaded = funcMidString(strUploaded, "datetime", "<sc")
    strUploaded = funcMidString(strUploaded, ">", "<")
    varFavorites = funcMidString(strData, "<i class=""fas fa-heart""></i><span>Favorite <span class=""nobold"">(", ")")
    strFullTitle = funcMidString(strData, "<h1 class=""title"">", "</h1>")
    strFullTitle = funcListMidString(strFullTitle, ">", "<")
    strFullTitle = replace(strFullTitle, vbCrlf, "")
    strCoverImage = funcMidString(strData, "<meta itemprop=""image"" content=""", """ />")

    funcNhentaiGetInfo = array(strTitle, strParodies, strCharacters, strTags, strArtists, strGroups, strLanguages, strCategories, varPages, strUploaded, varFavorites, fCode, strFullTitle, strCoverImage)
    '0 - Title
    '1 - Parodies
    '2 - Characters
    '3 - Tags
    '4 - Artists
    '5 - Groups
    '6 - Languages
    '7 - Categories
    '8 - Pages
    '9 - Uploaded
    '10 - Favorites
    '11 - ID
    '12 - Full title
    '13 - Cover Image
end function

class cNhentaiApi
    dim arrInfo

    public property let BookID(aCode)
        arrInfo = funcNhentaiGetInfo(aCode)
    end property

    public property get getTitle()
        getTitle = arrInfo(0)
    end property

    public property get getParodies()
        getParodies = arrInfo(1)
    end property

    public property get getCharacters()
        getCharacters = arrInfo(2)
    end property

    public property get getTags()
        getTags = arrInfo(3)
    end property

    public property get getArtists()
        getArtists = arrInfo(4)
    end property

    public property get getGroups()
        getGroups = arrInfo(5)
    end property

    public property get getLanguages()
        getLanguages = arrInfo(6)
    end property

    public property get getCategories()
        getCategories = arrInfo(7)
    end property

    public property get getPages()
        getPages = arrInfo(8)
    end property

    public property get getUploaded()
        getUploaded = arrInfo(9)
    end property

    public property get getFavorites()
        getFavorites = arrInfo(10)
    end property

    public property get getID()
        getID = arrInfo(11)
    end property

    public property get getFullTitle()
        getFullTitle = arrInfo(12)
    end property

    public property get getCoverURL()
        getCoverURL = arrInfo(13)
    end property
end class

class cNhentaiApiv2
    dim arrInfo, title, parodies, characters, tags, artists, groups, languages, categories, pages, uploaded, favorites, id, error, fullTitle, cover, exists

    function BookID(aCode)
        error = 0

        if aCode = id then
        else
            arrInfo = funcNhentaiGetInfo(aCode)

            if isarray(arrInfo) then
            else
                if arrInfo = 404 then
                    error = 404
                    title = ""
                    parodies = ""
                    characters = ""
                    tags = ""
                    artists = ""
                    groups = ""
                    languages = ""
                    categories = ""
                    pages = 0
                    uploaded = ""
                    favorites = 0
                    id = 0
                    fullTitle = ""
                    cover = ""
                    exists = false
                    set BookID = me
                    exit function
                end if
            end if
        end if

        title = arrInfo(0)
        parodies = arrInfo(1)
        characters = arrInfo(2)
        tags = arrInfo(3)
        artists = arrInfo(4)
        groups = arrInfo(5)
        languages = arrInfo(6)
        categories = arrInfo(7)
        pages = arrInfo(8)
        uploaded = arrInfo(9)
        favorites = arrInfo(10)
        id = arrInfo(11)
        fullTitle = arrInfo(12)
        cover = arrInfo(13)
        exists = true
        set BookID = me
    end function
end class

set api = new cNhentaiApi
set api2 = new cNhentaiApiv2

wscript.echo(api2.BookID(354125).cover)
wscript.echo(api2.BookID(177013).title)
api.BookID = 354125
wscript.echo(api.getTitle)
api2.BookID(177013)
wscript.echo(api2.artists)
wscript.echo(api2.tags)
wscript.echo api2.BookID(354125).artists