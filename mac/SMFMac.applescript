-- Excel calls fetchURL through AppleScriptTask. Cookies are anonymous and temporary.
on run argv
    if (count argv) is not 1 then return my envelope(0, "Expected METHOD, linefeed, URL")
    return my fetchURL(item 1 of argv)
end run

on fetchURL(requestText)
    set tempDirectory to ""
    try
        set separatorAt to offset of linefeed in requestText
        if separatorAt < 2 then error "Invalid request format"
        set requestMethod to text 1 thru (separatorAt - 1) of requestText
        set targetURL to text (separatorAt + 1) thru -1 of requestText
        if targetURL contains linefeed or targetURL contains return then error "A URL cannot contain a line break"
        if targetURL does not start with "https://" and targetURL does not start with "http://" then error "Only HTTP and HTTPS URLs are supported"
        if requestMethod is not in {"GET", "POST", "YAHOO"} then error "Unsupported request method"
        -- Type 4 also appears in non-Yahoo recipes. Those requests use GET.
        if requestMethod is "YAHOO" and my isYahooAPI(targetURL) then
            set tempDirectory to do shell script "/usr/bin/mktemp -d -t smf-mac"
            set reply to my yahooRequest(targetURL, tempDirectory & "/cookies.txt")
        else
            if requestMethod is "YAHOO" then set requestMethod to "GET"
            set reply to my curlRequest(targetURL, requestMethod, "")
        end if
        if tempDirectory is not "" then my removeTempDirectory(tempDirectory)
        return my envelope(item 1 of reply, item 2 of reply)
    on error messageText number errorNumber
        if tempDirectory is not "" then my removeTempDirectory(tempDirectory)
        return my envelope(0, "Request failed (" & errorNumber & "): " & messageText)
    end try
end fetchURL

on curlRequest(targetURL, requestMethod, cookieFile)
    set commandText to "/usr/bin/curl --http1.1 --silent --show-error --location --compressed --connect-timeout 8 --max-time 25 --proto '=http,https' --proto-redir '=http,https' --user-agent 'Mozilla/5.0 SMF-Mac/0.1' --write-out '\n%{http_code}'"
    set commandText to commandText & " --request " & quoted form of requestMethod
    if cookieFile is not "" then set commandText to commandText & " --cookie " & quoted form of cookieFile & " --cookie-jar " & quoted form of cookieFile
    set commandText to commandText & " --url " & quoted form of targetURL
    set responseText to do shell script commandText without altering line endings
    if (length of responseText) < 4 then error "curl returned an invalid response"
    set statusText to text -3 thru -1 of responseText
    set bodyText to ""
    if (length of responseText) > 4 then set bodyText to text 1 thru -5 of responseText
    return {statusText as integer, bodyText}
end curlRequest

on yahooRequest(targetURL, cookieFile)
    -- The cookie endpoint can return 404 while supplying the anonymous cookie.
    my curlRequest("https://fc.yahoo.com", "GET", cookieFile)
    set crumbReply to my curlRequest("https://query1.finance.yahoo.com/v1/test/getcrumb", "GET", cookieFile)
    if item 1 of crumbReply is not 200 then return crumbReply
    set crumbText to item 2 of crumbReply
    if crumbText is "" or (length of crumbText) > 100 then error "Yahoo returned an invalid crumb"
    if crumbText contains linefeed or crumbText contains return or crumbText contains "{" or crumbText contains "<" then error "Yahoo returned an invalid crumb"
    set authenticatedURL to my removeCrumb(targetURL)
    if authenticatedURL contains "?" then
        set authenticatedURL to authenticatedURL & "&crumb=" & my encodeQueryValue(crumbText)
    else
        set authenticatedURL to authenticatedURL & "?crumb=" & my encodeQueryValue(crumbText)
    end if
    return my curlRequest(authenticatedURL, "GET", cookieFile)
end yahooRequest

on isYahooAPI(targetURL)
    repeat with authority in {"https://query1.finance.yahoo.com/", "https://query2.finance.yahoo.com/"}
        if targetURL starts with (authority as text) then return true
    end repeat
    return false
end isYahooAPI

on removeCrumb(targetURL)
    set oldDelimiters to AppleScript's text item delimiters
    try
        set questionAt to offset of "?" in targetURL
        if questionAt is 0 then return targetURL
        set baseURL to text 1 thru (questionAt - 1) of targetURL
        set queryText to ""
        if questionAt < (length of targetURL) then set queryText to text (questionAt + 1) thru -1 of targetURL
        set AppleScript's text item delimiters to "&"
        set queryParts to text items of queryText
        set keptParts to {}
        repeat with queryPart in queryParts
            set queryPart to queryPart as text
            if queryPart does not start with "crumb=" and queryPart is not "" then set end of keptParts to queryPart
        end repeat
        set queryText to keptParts as text
        set AppleScript's text item delimiters to oldDelimiters
        if queryText is "" then return baseURL
        return baseURL & "?" & queryText
    on error messageText number errorNumber
        set AppleScript's text item delimiters to oldDelimiters
        error messageText number errorNumber
    end try
end removeCrumb

on encodeQueryValue(valueText)
    set encodedText to ""
    set hexDigits to "0123456789ABCDEF"
    repeat with aCharacter in characters of valueText
        set aCharacter to aCharacter as text
        considering case
            if "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~" contains aCharacter then
                set encodedText to encodedText & aCharacter
            else
                set charCode to ASCII number of aCharacter
                if charCode > 127 then error "Yahoo returned a non-ASCII crumb"
                set encodedText to encodedText & "%" & character ((charCode div 16) + 1) of hexDigits & character ((charCode mod 16) + 1) of hexDigits
            end if
        end considering
    end repeat
    return encodedText
end encodeQueryValue

on removeTempDirectory(directoryPath)
    try
        do shell script "/bin/rm -rf -- " & quoted form of directoryPath
    end try
end removeTempDirectory

on envelope(statusCode, bodyText)
    return "SMFHTTP/1" & linefeed & (statusCode as text) & linefeed & bodyText
end envelope

-- Excel's SaveAs interface can require HFS paths even when FullName returns POSIX.
on excelPath(posixPath)
    if posixPath does not start with "/" then error "Expected an absolute POSIX path"
    if posixPath contains linefeed or posixPath contains return then error "A path cannot contain a line break"
    return (POSIX file posixPath) as text
end excelPath

-- Stage reviewed build modules inside Excel's own container. This avoids the
-- folder-access dialog in Excel's external-file import path.
on stageModule(requestText)
    set parts to my splitFileRequest(requestText)
    set projectRoot to item 1 of parts
    set moduleName to item 2 of parts
    set homePath to POSIX path of (path to home folder)
    if homePath ends with "/Library/Containers/com.microsoft.Excel/Data/" then
        set containerHome to homePath
    else
        set containerHome to homePath & "Library/Containers/com.microsoft.Excel/Data/"
    end if
    set cacheFolder to containerHome & "Library/Caches/SMF-Mac-Build/v0.2"
    return my copyProjectModuleTo(projectRoot, moduleName, cacheFolder)
end stageModule

on copyProjectModuleTo(projectRoot, moduleName, destinationFolder)
    my validateAbsolutePath(projectRoot)
    my validateAbsolutePath(destinationFolder)
    my validateBaseName(moduleName)
    if moduleName does not end with ".bas" and moduleName does not end with ".cls" then error "Expected a VBA module or class file"
    set sourceFile to projectRoot & "/code/" & moduleName
    set destinationFile to destinationFolder & "/" & moduleName
    do shell script "/bin/mkdir -p " & quoted form of destinationFolder
    do shell script "/bin/cp -p " & quoted form of sourceFile & " " & quoted form of destinationFile
    return my excelPath(destinationFile)
end copyProjectModuleTo

-- Read element definitions through the same supported AppleScriptTask helper.
-- Missing numbered files are normal in SMF's 0..22 search.
on readElementFile(requestText)
    set parts to my splitFileRequest(requestText)
    set elementFolder to item 1 of parts
    set suffixText to item 2 of parts
    my validateBaseName(suffixText)
    set sourceFile to elementFolder & "/smf-elements-" & suffixText & ".txt"
    return my readOptionalUTF8File(sourceFile)
end readElementFile

on readElementLocation(elementFolder)
    my validateAbsolutePath(elementFolder)
    return my readOptionalUTF8File(elementFolder & "/smf-Elements-Location.txt")
end readElementLocation

on readOptionalUTF8File(sourceFile)
    try
        do shell script "/bin/test -f " & quoted form of sourceFile & " && /bin/test -s " & quoted form of sourceFile
    on error messageText number errorNumber
        if errorNumber is 1 then return ""
        error messageText number errorNumber
    end try
    return read (POSIX file sourceFile) as «class utf8»
end readOptionalUTF8File

on splitFileRequest(requestText)
    set separatorAt to offset of linefeed in requestText
    if separatorAt < 2 then error "Expected a folder, linefeed, and filename"
    set folderPath to text 1 thru (separatorAt - 1) of requestText
    set filenameText to text (separatorAt + 1) thru -1 of requestText
    my validateAbsolutePath(folderPath)
    my validateBaseName(filenameText)
    return {folderPath, filenameText}
end splitFileRequest

on validateAbsolutePath(pathText)
    if pathText does not start with "/" then error "Expected an absolute POSIX path"
    if pathText contains linefeed or pathText contains return then error "A path cannot contain a line break"
end validateAbsolutePath

on validateBaseName(nameText)
    if nameText is "" or nameText is "." or nameText is ".." then error "Expected a filename"
    repeat with aCharacter in characters of nameText
        considering case
            if "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_." does not contain (aCharacter as text) then error "Invalid filename character"
        end considering
    end repeat
end validateBaseName

-- Diagnostic logging never blocks a build. The path stays inside its project.
on logBuildState(requestText)
    try
        set separatorAt to offset of linefeed in requestText
        if separatorAt < 2 then return "unavailable"
        set projectRoot to text 1 thru (separatorAt - 1) of requestText
        set stateText to text (separatorAt + 1) thru -1 of requestText
        my validateAbsolutePath(projectRoot)
        if stateText contains linefeed or stateText contains return then return "unavailable"
        set logFile to projectRoot & "/mac/native-build-state.log"
        do shell script "/usr/bin/printf '%s\\n' " & quoted form of stateText & " >> " & quoted form of logFile
        return "logged"
    on error
        return "unavailable"
    end try
end logBuildState

-- Comment fills are not supported by current Excel for Mac. Images are
-- downloaded as files for Shapes.AddPicture, without passing binary data
-- through an AppleScriptTask text result.
on downloadImage(targetURL)
    set imageDirectory to ""
    set responseStatus to 0
    try
        if targetURL contains linefeed or targetURL contains return then error "An image URL cannot contain a line break"
        if targetURL does not start with "https://" and targetURL does not start with "http://" then error "Only HTTP and HTTPS image URLs are supported"
        set cacheFolder to my imageCacheFolder()
        do shell script "umask 077; /bin/mkdir -p " & quoted form of cacheFolder
        set imageDirectory to do shell script "/usr/bin/mktemp -d " & quoted form of (cacheFolder & "/image.XXXXXX")
        do shell script "/usr/bin/printf '%s' 'SMFIMAGE/1' > " & quoted form of (imageDirectory & "/.smf-image-owner")
        set bodyFile to imageDirectory & "/body.bin"
        set imageFile to imageDirectory & "/image.png"
        set commandText to "/usr/bin/curl --globoff --http1.1 --silent --show-error --location --compressed --connect-timeout 8 --max-time 25 --max-filesize 10485760 --proto '=http,https' --proto-redir '=http,https' --user-agent 'Mozilla/5.0 SMF-Mac/0.4' --write-out '%{http_code}'"
        set commandText to commandText & " --output " & quoted form of bodyFile & " --url " & quoted form of targetURL
        set responseStatus to (do shell script commandText) as integer
        if responseStatus is not 200 then error "Image request returned HTTP " & responseStatus
        my decodeImageFile(bodyFile, imageFile)
        return my imageEnvelope(200, imageFile, "")
    on error messageText number errorNumber
        if imageDirectory is not "" then my releaseImage(imageDirectory & "/image.png")
        return my imageEnvelope(responseStatus, "", "Image request failed (" & errorNumber & "): " & messageText)
    end try
end downloadImage

-- Local pictures also pass through Excel's own cache. Excel can reject a
-- direct path outside its container even when AppleScript can read the file.
on prepareLocalImage(sourceFile)
    set imageDirectory to ""
    try
        my validateAbsolutePath(sourceFile)
        do shell script "/bin/test -f " & quoted form of sourceFile
        set byteCount to (do shell script "/usr/bin/wc -c < " & quoted form of sourceFile) as integer
        if byteCount < 1 or byteCount > 10485760 then error "Image size must be between 1 byte and 10 MiB"
        set cacheFolder to my imageCacheFolder()
        do shell script "umask 077; /bin/mkdir -p " & quoted form of cacheFolder
        set imageDirectory to do shell script "/usr/bin/mktemp -d " & quoted form of (cacheFolder & "/image.XXXXXX")
        do shell script "/usr/bin/printf '%s' 'SMFIMAGE/1' > " & quoted form of (imageDirectory & "/.smf-image-owner")
        set bodyFile to imageDirectory & "/body.bin"
        set imageFile to imageDirectory & "/image.png"
        do shell script "/bin/cp -p " & quoted form of sourceFile & " " & quoted form of bodyFile
        my decodeImageFile(bodyFile, imageFile)
        return my imageEnvelope(200, imageFile, "")
    on error messageText number errorNumber
        if imageDirectory is not "" then my releaseImage(imageDirectory & "/image.png")
        return my imageEnvelope(0, "", "Local image failed (" & errorNumber & "): " & messageText)
    end try
end prepareLocalImage

on decodeImageFile(bodyFile, imageFile)
    set byteCount to (do shell script "/usr/bin/wc -c < " & quoted form of bodyFile) as integer
    if byteCount < 1 or byteCount > 10485760 then error "Image size must be between 1 byte and 10 MiB"
    set mimeType to do shell script "/usr/bin/file --brief --mime-type " & quoted form of bodyFile
    if mimeType is not in {"image/png", "image/jpeg", "image/gif"} then error "Response is not a PNG, JPEG, or GIF image"
    set dimensions to do shell script "/usr/bin/sips -g pixelWidth -g pixelHeight " & quoted form of bodyFile
    set pixelWidth to my imageDimension(dimensions, "pixelWidth: ")
    set pixelHeight to my imageDimension(dimensions, "pixelHeight: ")
    if pixelWidth < 1 or pixelHeight < 1 or pixelWidth > 8192 or pixelHeight > 8192 or pixelWidth * pixelHeight > 16777216 then error "Image dimensions exceed the supported limit"
    -- Decode before Excel sees the file. This rejects corrupt image bodies
    -- and gives AddPicture one supported, static PNG representation.
    do shell script "/usr/bin/sips -s format png " & quoted form of bodyFile & " --out " & quoted form of imageFile
    do shell script "/bin/test -s " & quoted form of imageFile & " && /bin/rm -f -- " & quoted form of bodyFile
end decodeImageFile

on imageCacheFolder()
    set homePath to POSIX path of (path to home folder)
    if homePath ends with "/Library/Containers/com.microsoft.Excel/Data/" then
        set containerHome to homePath
    else
        set containerHome to homePath & "Library/Containers/com.microsoft.Excel/Data/"
    end if
    return containerHome & "Library/Caches/SMF-Mac-Images"
end imageCacheFolder

on imageDimension(dimensionText, labelText)
    repeat with dimensionLine in paragraphs of dimensionText
        set dimensionLine to dimensionLine as text
        set labelAt to offset of labelText in dimensionLine
        if labelAt > 0 then return (text (labelAt + (length of labelText)) thru -1 of dimensionLine) as integer
    end repeat
    error "Image decoder did not return its dimensions"
end imageDimension

on imageEnvelope(statusCode, imagePath, errorText)
    set errorText to my replaceImageLineBreaks(errorText)
    return "SMFIMAGE/1" & linefeed & (statusCode as text) & linefeed & imagePath & linefeed & errorText
end imageEnvelope

on replaceImageLineBreaks(valueText)
    set oldDelimiters to AppleScript's text item delimiters
    try
        set AppleScript's text item delimiters to {return, linefeed}
        set pieces to text items of valueText
        set AppleScript's text item delimiters to " "
        set valueText to pieces as text
        set AppleScript's text item delimiters to oldDelimiters
        return valueText
    on error messageText number errorNumber
        set AppleScript's text item delimiters to oldDelimiters
        error messageText number errorNumber
    end try
end replaceImageLineBreaks

on releaseImage(imagePath)
    try
        set cacheFolder to my imageCacheFolder()
        set expectedPrefix to cacheFolder & "/image."
        if imagePath does not start with expectedPrefix or imagePath does not end with "/image.png" then return "ignored"
        set directoryPart to text ((length of expectedPrefix) + 1) thru ((length of imagePath) - 10) of imagePath
        if directoryPart is "" then return "ignored"
        repeat with aCharacter in characters of directoryPart
            considering case
                if "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789" does not contain (aCharacter as text) then return "ignored"
            end considering
        end repeat
        set imageDirectory to expectedPrefix & directoryPart
        set ownerFile to imageDirectory & "/.smf-image-owner"
        if (do shell script "/bin/cat " & quoted form of ownerFile) is not "SMFIMAGE/1" then return "ignored"
        -- Delete only this request's fixed filenames. rmdir refuses a directory
        -- containing any other file; no caller-provided path reaches rm -rf.
        do shell script "/bin/rm -f -- " & quoted form of (imageDirectory & "/image.png") & " " & quoted form of (imageDirectory & "/body.bin") & " " & quoted form of ownerFile
        do shell script "/bin/rmdir -- " & quoted form of imageDirectory
        return "removed"
    on error
        return "unavailable"
    end try
end releaseImage
