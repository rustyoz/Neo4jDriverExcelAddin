# Installer for the Addin

This is a WiX project to create an MSI installer to install / uninstall the addin to Excel.

## Requirements

You'll need the WiX toolset from: https://wixtoolset.org/releases/
This has been tested/compiled with `3.11.2` later versions might need changes.

## How do I get an installer?

I generally use the command line to create the installer - with the following code:

```
rm *.msi
rm *.wixpdb
rm *.wixobj

& 'LOCATION_OF_WIX\wix\candle.exe' .\Product.wxs -dConfiguration="debug"
& 'LOCATION_OF_WIX\wix\light.exe' -out Install.Neo4j.Excel.Addin.msi -ext WixUIExtension .\Product.wixobj
```

## Notes

* Candle.exe creates `.wixobj` - I would run the `rm` methods first to ensure you don't end up rebuilding older versions.
* `-dConfiguration="debug"` is a property that is in the `product.wxs` file - you can see in the `Component\File` sections (`Source="..\Neo4jDriverExcelAddin\bin\$(var.Configuration)\Neo4jDriverExcelAddin.vsto"`) - change to `Release` for release builds... ETC :)

## Debugging

Generally you're looking at using log files - to generate one from the installer run the following:

```
msiexec /i Install.Neo4j.Excel.Addin.msi /l*v installer.log
```

## Issues?

* License.rtf - This is just the MIT license - ff you edit this - you **need** to use WordPad to save it. Nothing else will work. 