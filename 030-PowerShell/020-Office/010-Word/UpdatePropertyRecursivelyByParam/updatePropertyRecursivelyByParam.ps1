# In this example, It would update Document Property, Keywords in Testing.docx in the same folder.
# 	The original value of Keywords is Testing 123. 
# 	After executing this program, the field would be updated to Testing Updated

param (
	[string]$folder = ".",
	[string]$property = "Keywords",
	[string]$value = "Updated",
	[string]$printOrginalValue = $false,
	[string]$wordVisible = $false
)

$docxs = Get-ChildItem $folder -Recurse -Include *docx;
foreach ($docx in $docxs){
	#Create Word application
	$word = New-Object -ComObject word.application;

	if ($wordVisible.ToUpper() -eq "TRUE"){
		$word.Visible = $True
	}

	#Get the reference to the document
	$document = $word.documents.open($docx.fullname);

	#set up binding flags for built properties
	$binding = "System.Reflection.BindingFlags" -as [type];        
	$builtProperties = $document.BuiltInDocumentProperties;
	$builtPropertiesType = $builtProperties.GetType()

	[Array]$propertyName = $property
	[Array]$propertyValue = $value

	$myProperty = $builtPropertiesType.InvokeMember("Item", $binding::GetProperty, $null, $builtProperties, $propertyName)
	if ($printOrginalValue.ToUpper() -eq "TRUE"){
		$orginalPropertyValue  = $builtPropertiesType.InvokeMember("value",$binding::GetProperty,$null,$myProperty,$null);
		write-host ("origianl property value: " + $orginalPropertyValue)
	}
	
	#Set property value
	$builtPropertiesType.InvokeMember("Value",$binding::SetProperty,$null,$myProperty,$propertyValue)

	$document.Fields.Update() | Out-Null
	foreach ($Section in $document.Sections)
	{
		## Update Header
		#$Header = $Section.Headers.Item(1)
		#$Header.Range.Fields.Update() | Out-Null
    
		# Update Footer
		$Footer = $Section.Footers.Item(1)
		$Footer.Range.Fields.Update() | Out-Null
	}

	"[Updated] " + $docx.fullname
	
	#save and close document
	$document.Saved = $false;
	$document.save();
	$word.Quit();

}