# In this example, It would update Document Property, Keywords in Testing.docx in the same folder.
# 	The original value of Keywords is Testing 123. 
# 	After executing this program, the field would be updated to Testing Updated

#Create Word application
$word = New-Object -ComObject word.application;

#Get the reference to the document
$document = $word.documents.open($(get-location).Path + "\Testing.docx");

#set up binding flags for built properties
$binding = "System.Reflection.BindingFlags" -as [type];        
$builtProperties = $document.BuiltInDocumentProperties;
$builtPropertiesType = $builtProperties.GetType()

[Array]$propertyName = "Keywords"
[Array]$propertyValue = "Testing Updated"

$myProperty = $builtPropertiesType.InvokeMember("Item", $binding::GetProperty, $null, $builtProperties, $propertyName)
$myPropertyValue  = [System.__ComObject].InvokeMember("value",$binding::GetProperty,$null,$myProperty,$null);
write-host $myPropertyValue

#Set property value
$builtPropertiesType.InvokeMember("Value",$binding::SetProperty,$null,$myProperty,$propertyValue)

$document.Fields.Update() | Out-Null
foreach ($Section in $document.Sections)
{
	# Update Header
	$Header = $Section.Headers.Item(1)
	$Header.Range.Fields.Update()

	# Update Footer
	$Footer = $Section.Footers.Item(1)
	$Footer.Range.Fields.Update()
}

#save and close document
$document.Saved = $false;
$document.save();
$word.Quit();
