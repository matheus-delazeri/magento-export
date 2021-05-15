# Magento 1.9 Export products module
A magento 1.9 module to export products informations to a .XLSX file.

## Module informations
`Package/Namespace`: "Matheus"  

`Modulename`: "Export"

`codepool`: "community"  

## How to install
Add the folder `Matheus` inside `/app/code/community/` and add the file `Matheus_Export.xml` inside `/app/etc/modules/`

## How to use
After installation a new submenu named `Export Products` will be created at the menu `Catalog` in your admin panel. Click in it to enter the module's page. 

Now, you just need to press the button `Export` to generate a .XLSX file, that will be automatically downloaded in your browser, containing the meaningful informations about the products in your store.

## The .XLSX file
The sheet generated by the module will be split in 2 different pages:
* `All products`
* `Association`

The page `All products` will have all the info about the products:
|store|websites|attribute_set|type|categories|sku|name|price|special_price|weight|status|visibility|tax_class_id|description|short_description|qty|is_in_stock|store_id|mgs_brand|country_of_manufacture|
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
|admin|base|default|configurable|Default Category|product-sku|Product name|100|60|1|1|4|none|description|short description|5|1|0||brasil|

Obs.: custom attributes (like color, size, etc) will be placed after the last column in the page `All products`

And the page `Association` will have the info about the association between configurable products and their children:
|sku|_super_products_sku|_super_attribute_code|_super_attribute_option|
| --- | --- | --- | --- |
|configurable-sku||size||
||child-product-P|size|P|
||child-product-M|size|M|

You can also check the `Example Sheet.xlsx` to see which infos are stored.
