# Create object to access the scanner
$deviceManager = new-object -ComObject WIA.DeviceManager
$device = $deviceManager.DeviceInfos.Item(1).Connect()

# Create object to access the scanned image later
$imageProcess = new-object -ComObject WIA.ImageProcess

# Store file format GUID strings
$wiaFormatBMP  = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatPNG  = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatGIF  = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatJPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatTIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"

$device.Properties.Item("3088") = 2
$device.Properties.Item("3096") = 1

$error.Clear()

foreach ($item in $device.Items) {
    $item.Properties.Item("6146") = 1
    $item.Properties.Item("6151") = 1102
    $item.Properties.Item("6152") = 1700
    $image = $item.Transfer() 
}


# set type to JPEG and quality/compression level
$imageProcess.Filters.Add($imageProcess.FilterInfos.Item("Convert").FilterID)
$imageProcess.Filters.Item($imageProcess.Filters.Count).Properties.Item("FormatID").Value = $wiaFormatJPEG
$image = $imageProcess.Apply($image)

# Build filepath from desktop path and filename 'scan0'
$filename = "$([Environment]::GetFolderPath("Desktop"))\scan{0}.jpg"

# If a file named 'scan0' already exists, increment the index as long as needed
$index = 0
while (test-path ($filename -f $index)) {
    [void](++$index)
}
$filename = $filename -f $index

# Save image to 'C:\Users\<username>\Desktop\scan{x}'
$image.SaveFile($filename)
