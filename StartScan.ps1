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

$device.Properties.Item("3088") = 1
$device.Properties.Item("3096") = 1

$error.Clear()

foreach ($item in $device.Items) {
    $item.Properties.Item("6146") = 4
    $item.Properties.Item("6151") = 1654
    $item.Properties.Item("6152") = 2334
    $image = $item.Transfer() 
}


while(-not $error) {
    try {
        foreach ($item in $device.Items) {
            $item.Properties.Item("6146") = 4
            $item.Properties.Item("6151") = 1654
            $item.Properties.Item("6152") = 2334
            $frame = $item.Transfer() 
            $imageProcess.Filters.Add($imageProcess.FilterInfos.Item("Frame").FilterID)
            $imageProcess.Filters.Item($imageProcess.Filters.Count).Properties.Item("ImageFile") = $frame;
        }
    }
    catch {
    }
}

# set type to JPEG and quality/compression level
$imageProcess.Filters.Add($imageProcess.FilterInfos.Item("Convert").FilterID)
$imageProcess.Filters.Item($imageProcess.Filters.Count).Properties.Item("FormatID").Value = $wiaFormatTIFF
$image = $imageProcess.Apply($image)

# Build filepath from desktop path and filename 'Scan 0'
$filename = "$([Environment]::GetFolderPath("Desktop"))\Scan {0}.tif"

# If a file named 'Scan 0' already exists, increment the index as long as needed
$index = 0
while (test-path ($filename -f $index)) {
    [void](++$index)
}
$filename = $filename -f $index

# Save image to 'C:\Users\<username>\Desktop\Scan {x}'
$image.SaveFile($filename)