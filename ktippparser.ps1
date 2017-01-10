$OutPath = "f:\"
$PhoneBookName = "CallCenterPhonebook"
$MaxPhoneBookEntries = 1500 #Value where a new Phonebook XML file will be created.
$CurrPage = 1
$CurrPhoneBook = 1
$BaseUrl = "https://www.ktipp.ch/service/warnlisten/detail/w/unerwuenschte-oder-laestige-telefonanrufe/?keyword=&ajax=ajax-search-form&page="
$page = Invoke-WebRequest -Uri "$BaseUrl$CurrPage"
#Get the total Entries from the top of the page
$TotalEntries = [regex]::Matches($page.RawContent,"<i>(.*?) Einträge").Groups[1].Value
$EntriesPerPage = 100
$Pages = [math]::Round($TotalEntries / $EntriesPerPage,0)
#$Pages = 15
$currNumber = 1

function CreateBlankXML()
{
    [xml]$Xml = Get-Content -Path C:\Users\marcl\Documents\ktipp_template.xml
    $xml
}

$xml = CreateBlankXML
$contacts = $Xml.phonebooks.phonebook.contact

Write-Host -ForegroundColor Green "Ktipp Sperrlisten Generator. Einträge: $TotalEntries"

while($CurrPage -le $Pages)
{
    Write-Host -ForegroundColor Green "Processing Page $CurrPage/$Pages"
    $page = Invoke-WebRequest -Uri "$BaseUrl$CurrPage"
    $regex = [regex]::Matches($page,"<strong>(.*?)</strong>")
    $xml.phonebooks.phonebook.name = "$PhoneBookName $CurrPhoneBook"
    $PhonebookFileName = "$OutPath$PhoneBookName-$CurrPhoneBook.xml"

    foreach($group in $regex.Groups){
        #We filter all found regex groups starting with <
        #We filter all found regex groups where the telephone number also contains strings
        #Some numbers have some comments in it.. too bad..
        if((-not $group.Value.StartsWith("<")) -and ($group.Value -notmatch "[a-z(?/,;-A-Z]"))
        {
            #Add every phone number to the phone book
            $contact = $contacts[0].Clone()
            $contact.person.realName = "#$($currNumber.ToString("0000")) - $($group.Value)"
            $contact.telephony.number = $group.Value
            $xml.phonebooks.phonebook.AppendChild($contact)
            Write-Host -ForegroundColor Green "Added $($group.Value) to the phonebook as entry $currNumber"
            $currNumber += 1

            #Create new XML File
            if($CurrNumber % $MaxPhoneBookEntries -eq 0)
            {
                Write-Host -ForegroundColor Green "Saving $PhonebookFileName"
                $xml.Save($PhonebookFileName)
                $xml = CreateBlankXML
                $contacts = $Xml.phonebooks.phonebook.contact
                $CurrPhoneBook += 1
            }

        }
    }
    $Currpage += 1
}
Write-Host -ForegroundColor Green "Saving $PhonebookFileName"
$xml.Save($PhonebookFileName)
Write-Host -ForegroundColor Green "Filtered out $($TotalEntries-$currNumber) of $TotalEntries entries due to some wrong telephone numbers"