class NhentaiApi
    dim title, parodies, characters, tags, artists, groups, languages, categories, pages, uploaded, favorites, id, error, fullTitle, cover, statusCode
    private objShell, objFile, objHTML, objHTTP, strDirectory, strData

    private sub Class_Initialize
        set objShell = CreateObject("wscript.shell")
        set objFile = CreateObject("Scripting.FileSystemObject")
        set objHTML = CreateObject("htmlfile")
        set objHTTP = CreateObject("MSXML2.XMLHTTP")
        strDirectory = objShell.CurrentDirectory
    end sub

    private function funcMidString(fInput, fLeft, fRight)
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

    private function funcListMidString(fInput, fLeft, fRight)
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

    private sub sbSetMultipleVariables(fVariables, fValue)
        dim arrTemp
        arrTemp = split(fVariables, ",")

        for each element in arrTemp
            execute(element & " = " & fValue)
        next
    end sub

    function BookID(code)
        set BookID = me
        objHTTP.open "GET", "https://nhentai.net/g/" & code, false
        objHTTP.send
        strData = objHTTP.responsetext
        objFile.createtextfile(strDirectory & "/" & fCode & ".html", true, true).writeline(strData)

        statusCode = objHTTP.status

        if statusCode = 404 then
            call sbSetMultipleVariables("title, parodies, characters, tags, artists, groups, languages, categories, pages, uploaded, favorites, id, error, fullTitle, cover", false)
            exit function
        end if

        title = funcMidString(strData, "<meta itemprop=""name"" content=""", """ />")
        parodies = funcMidString(strData, "Parodies:", "Characters:")
        parodies = funcListMidString(parodies, "<span class=""name"">", "</span>")
        characters = funcMidString(strData, "Characters:", "Tags:")
        characters = funcListMidString(characters, "<span class=""name"">", "</span>")
        'tags = funcMidString(strData, "Tags:", "Artists:")
        'tags = funcListMidString(tags, "<span class=""name"">", "</span>")
        tags = funcMidString(strData, "<meta name=""twitter:description"" content=""", """ />")
        tags = replace(tags, ", ", vbCrlf)
        artists = funcMidString(strData, "Artists:", "Groups:")
        artists = funcListMidString(artists, "<span class=""name"">", "</span>")
        groups = funcMidString(strData, "Groups:", "Languages:")
        groups = funcListMidString(groups, "<span class=""name"">", "</span>")
        languages = funcMidString(strData, "Languages:", "Categories:")
        languages = funcListMidString(languages, "<span class=""name"">", "</span>")
        categories = funcMidString(strData, "Categories:", "Pages:")
        categories = funcListMidString(categories, "<span class=""name"">", "</span>")
        pages = funcMidString(strData, "Pages:", "Uploaded:")
        pages = funcListMidString(pages, "<span class=""name"">", "</span>")
        uploaded = funcMidString(strData, "Uploaded:", "ript>")
        uploaded = funcMidString(uploaded, "datetime", "<sc")
        uploaded = funcMidString(uploaded, ">", "<")
        favorites = funcMidString(strData, "<i class=""fas fa-heart""></i><span>Favorite <span class=""nobold"">(", ")")
        fullTitle = funcMidString(strData, "<h1 class=""title"">", "</h1>")
        fullTitle = funcListMidString(fullTitle, ">", "<")
        fullTitle = replace(fullTitle, vbCrlf, "")
        cover = funcMidString(strData, "<meta itemprop=""image"" content=""", """ />")
    end function
end class

set api = new NhentaiApi

api.BookID(177013)
wscript.echo(api.BookID(215600).statusCode)
wscript.echo(api.title)
