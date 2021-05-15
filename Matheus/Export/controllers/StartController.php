<?php

class Matheus_Export_StartController extends Mage_Adminhtml_Controller_Action{
	public function indexAction() {
		/** Include PHPExcel */
		require_once dirname(__FILE__).'/../Classes/PHPExcel.php';
		/** Create new PHPExcel object */
		$objPHPExcel = new PHPExcel();
		/** Get all products in the store */
		$allProducts = Mage::getModel('catalog/product')->getCollection()				      
						  ->addAttributeToSelect('*');
		$row_all = 2;
		$row_assoc = 2;
		$this->setHeaderNames($objPHPExcel);
		foreach ($allProducts as $product){
			$product = Mage::getModel('catalog/product')->load($product->getId());
			$objPHPExcel->setActiveSheetIndex(0)
				/** Set constants */
				->setCellValueByColumnAndRow(0, $row_all, "admin")   //store
				->setCellValueByColumnAndRow(1, $row_all, "base")    //website
				->setCellValueByColumnAndRow(2, $row_all, "default")  //attribute_set
				->setCellValueByColumnAndRow(12, $row_all, "none")   //tax_class_id
				->setCellValueByColumnAndRow(17, $row_all, "0")      //store_id
				->setCellValueByColumnAndRow(18, $row_all, "")       //mgs_brand
				->setCellValueByColumnAndRow(19, $row_all, "brasil") //country_of_manufacture
				->setCellValueByColumnAndRow(20, $row_all, "0")      //leadtime
				/** Set variables */
				->setCellValueByColumnAndRow(3, $row_all, $product->getTypeId())                    //product_type
				->setCellValueByColumnAndRow(5, $row_all, $product->getSku())                       //sku
				->setCellValueByColumnAndRow(6, $row_all, $product->getName())                      //name
				->setCellValueByColumnAndRow(7, $row_all, $product->getPrice())                     //price
				->setCellValueByColumnAndRow(8, $row_all, $product->getSpecialPrice())             //special_price
				->setCellValueByColumnAndRow(9, $row_all, $product->getWeight())                   //weight
				->setCellValueByColumnAndRow(10, $row_all, $product->getStatus())                   //status
				->setCellValueByColumnAndRow(11, $row_all, $product->getVisibility())               //visibility
				->setCellValueByColumnAndRow(13, $row_all, $product->getDescription())              //description
				->setCellValueByColumnAndRow(14, $row_all, $product->getShortDescription());        //short_description
			/** Set children */
			if ($product->getTypeId() == 'configurable'){
				$children = $product->getTypeInstance()->getUsedProducts($product->getId());
				if(!empty($children)){
					$row_assoc = $this->setChildrenProducts($objPHPExcel, $children, $product, $row_assoc);
				}
			} 
			/** Set categories */
			if (!empty($product->getCategoryids())){
				$this->setCategoryNames($objPHPExcel, $product->getCategoryIds(), $row_all);
			}
			/** Set stock fields */
			$this->setStockFields($objPHPExcel, $product, $row_all);
			/** Set attributes */
			$this->setSpecificAttributes($objPHPExcel, $product, $row_all);
			$row_all += 1;
		}
		/** Save Excel 2007 file */
		try{
			$objPHPExcel->setActiveSheetIndex(1)->setTitle("Association");
			$objPHPExcel->setActiveSheetIndex(0)->setTitle("All products");
			$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
			header('Content-type: application/vnd.ms-excel');
			header('Content-Disposition: attachment; filename="all_products.xlsx"');
			$objWriter->save('php://output');
		} catch(Exception $e){
			throw new Exception('Error while saving Excel file.');
		}	
	}
	private function setHeaderNames($objPHPExcel){
		$objPHPExcel->setActiveSheetIndex(0)
			->setCellValueByColumnAndRow(0, 1, "store")
			->setCellValueByColumnAndRow(1, 1, "websites")
			->setCellValueByColumnAndRow(2, 1, "attribute_set")
			->setCellValueByColumnAndRow(3, 1, "type")
			->setCellValueByColumnAndRow(4, 1, "categories")
			->setCellValueByColumnAndRow(5, 1, "sku")
			->setCellValueByColumnAndRow(6, 1, "name")
			->setCellValueByColumnAndRow(7, 1, "price")
			->setCellValueByColumnAndRow(8, 1, "special_price")
			->setCellValueByColumnAndRow(9, 1, "weight")
			->setCellValueByColumnAndRow(10, 1, "status")
			->setCellValueByColumnAndRow(11, 1, "visibility")
			->setCellValueByColumnAndRow(12, 1, "tax_class_id")
			->setCellValueByColumnAndRow(13, 1, "description")
			->setCellValueByColumnAndRow(14, 1, "short_description")
			->setCellValueByColumnAndRow(15, 1, "qty")
			->setCellValueByColumnAndRow(16, 1, "is_in_stock")
			->setCellValueByColumnAndRow(17, 1, "store_id")
			->setCellValueByColumnAndRow(18, 1, "mgs_brand")
			->setCellValueByColumnAndRow(19, 1, "country_of_manufacture")
			->setCellValueByColumnAndRow(20, 1, "leadtime");

		$objPHPExcel->createSheet(1);
		$objPHPExcel->setActiveSheetIndex(1)
			->setCellValueByColumnAndRow(0, 1, "sku")
			->setCellValueByColumnAndRow(1, 1, "_super_products_sku")
			->setCellValueByColumnAndRow(2, 1, "_super_attribute_code")
			->setCellValueByColumnAndRow(3, 1, "_super_attribute_option");
	}

	private function setChildrenProducts($objPHPExcel, $children, $product, $row_assoc){
		Mage::app()->setCurrentStore(Mage_Core_Model_App::ADMIN_STORE_ID);
		$productAttributeOptions = $product->getTypeInstance(true)->getConfigurableAttributesAsArray($product);
		$objPHPExcel->setActiveSheetIndex(1)
			->setCellValueByColumnAndRow(0, $row_assoc, $product->getSku());
		foreach($productAttributeOptions as $productAttribute){
			$atr_code = $productAttribute['attribute_code'];
			$objPHPExcel->setActiveSheetIndex(1)
				->setCellValueByColumnAndRow(0, $row_assoc, $product->getSku());
			$objPHPExcel->setActiveSheetIndex(1)
				->setCellValueByColumnAndRow(2, $row_assoc, $atr_code);
			$row_assoc+=1;
			foreach($children as $child){
				$child = Mage::getModel('catalog/product')->loadByAttribute('sku', $child['sku']);
				$objPHPExcel->setActiveSheetIndex(1)
					->setCellValueByColumnAndRow(1, $row_assoc, $child->getSku());
				$objPHPExcel->setActiveSheetIndex(1)
					->setCellValueByColumnAndRow(2, $row_assoc, $atr_code);
				$objPHPExcel->setActiveSheetIndex(1)
					->setCellValueByColumnAndRow(3, $row_assoc, $child->getAttributeText($atr_code));
				$row_assoc+=1;
			}
		}
		return $row_assoc;
	}

	private function setCategoryNames($objPHPExcel, $categories, $row_all){
		$categoriesNames = '';
		foreach ($categories as $category){
			$cat = Mage::getModel('catalog/category')
				->setStoreId(Mage::app()->getStore()->getId())->load($category);
		$categoriesNames .= $cat->getName().',';
		}
		$categoriesNames = substr($categoriesNames,0,-1);
		$objPHPExcel->setActiveSheetIndex(0)
			->setCellValueByColumnAndRow(4, $row_all, $categoriesNames);
	}

	private function setSpecificAttributes($objPHPExcel, $product, $row_all) {
    $defaultAttributes = array("name","sku","description","short_description","old_id","weight","news_from_date","news_to_date","url_path","status","url_key","category_ids",
	"visibility","country_of_manufacture","required_options","has_options","image_label","small_image_label","thumbnail_label","created_at","updated_at","price_type","sku_type",
	"weight_type","shipment_type","links_purchased_separately","samples_title","links_title","links_exist","price","group_price","special_price","special_from_date","special_to_date",
	"cost","tier_price","minimal_price","msrp_enabled","msrp_display_actual_price_type","msrp","tax_class_id","price_view","meta_title","meta_keyword","meta_description","is_recurring",
	"recurring_profile","custom_design","custom_design_from","custom_design_to","custom_layout_update","page_layout","options_container","gift_message_available");

    $attributeSetId = $product->getAttributeSetId();
    $allAttributes = Mage::getModel('catalog/product_attribute_api')->items($attributeSetId);
    $attributeCodes = array();
    foreach ($allAttributes as $attribute) {
        $attributeCodes[] = $attribute['code'];
    }
    $currentAttributes = array_diff($attributeCodes, $defaultAttributes);
	$index = 0;
	foreach ($currentAttributes as $attributeCode){
		$objPHPExcel->setActiveSheetIndex(0)
			->setCellValueByColumnAndRow(20+$index, 1, $attributeCode);
		if(!empty($product->getAttributeText($attributeCode))){
			$objPHPExcel->setActiveSheetIndex(0)
				->setCellValueByColumnAndRow(20+$index, $row_all, $product->getAttributeText($attributeCode));
		} elseif(!empty($product->getData($attributeCode))){
			$objPHPExcel->setActiveSheetIndex(0)
				->setCellValueByColumnAndRow(20+$index, $row_all, $product->getData($attributeCode));
		}
		$index+=1;
	}
    }

    private function setStockFields($objPHPExcel, $product, $row_all){
		$stock = Mage::getModel('cataloginventory/stock_item')->loadByProduct($product);
		$in_stock = 0;
		if ($stock->getQty() > 0) {
			$in_stock = 1;	
		}
		$objPHPExcel->setActiveSheetIndex(0)
			->setCellValueByColumnAndRow(15, $row_all, $stock->getQty())
			->setCellValueByColumnAndRow(16, $row_all, $in_stock);
    }
}