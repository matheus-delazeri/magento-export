<?php

class Matheus_Export_StartController extends Mage_Adminhtml_Controller_Action{
	public function indexAction() {
		/** Export file format */
		$fileFormat = $this->getRequest()->getPost('file_type');
		/** Date range for export */
		$date_start = $this->getRequest()->getPost('date_start');
		$date_end = $this->getRequest()->getPost('date_end');
		/** Pages to export */
		$page = $this->getRequest()->getPost('sheets');
		$assoc_index = 1;
		if($page=='assoc'){$assoc_index=0;};
		/** Chosen category */
		$category = $this->getRequest()->getPost('categories');
		/** Include PHPExcel */
		require_once dirname(__FILE__).'/../Classes/PHPExcel.php';
		/** Create new PHPExcel object */
		$objPHPExcel = new PHPExcel();
		/** Get all products in the store */
		$allProducts = Mage::getModel('catalog/product')->getCollection()				      
						  ->addAttributeToSelect('*');
		$row_all = 2;
		$row_assoc = 2;
		$this->setHeaderNames($objPHPExcel, $page, $assoc_index);
		foreach ($allProducts as $product){
			$product = Mage::getModel('catalog/product')->load($product->getId());
			$atrSet = Mage::getModel("eav/entity_attribute_set");
			$atrSet->load($product->getAttributeSetId());
			$product_date = substr($product->getData('created_at'), 0, -15);

			$inChosenCat = $this->checkIfInChosenCat($product->getCategoryIds(), $category);
			$inDateRange = $this->checkIfInDateRange($product_date, $date_start, $date_end);
			if($inDateRange && $inChosenCat){
				if($page!='assoc'){
					$objPHPExcel->setActiveSheetIndex(0)
						/** Set constants */
						->setCellValueByColumnAndRow(0, $row_all, "admin")   //store
						->setCellValueByColumnAndRow(1, $row_all, "base")    //website
						->setCellValueByColumnAndRow(12, $row_all, "none")   //tax_class_id
						->setCellValueByColumnAndRow(17, $row_all, "0")      //store_id
						/** Set variables */
						->setCellValueByColumnAndRow(2, $row_all, $atrSet->getAttributeSetName())   //attribute_set
						->setCellValueByColumnAndRow(3, $row_all, $product->getTypeId())            //product_type
						->setCellValueByColumnAndRow(5, $row_all, $product->getSku())               //sku
						->setCellValueByColumnAndRow(6, $row_all, $product->getName())              //name
						->setCellValueByColumnAndRow(7, $row_all, $product->getPrice())             //price
						->setCellValueByColumnAndRow(8, $row_all, $product->getSpecialPrice())      //special_price
						->setCellValueByColumnAndRow(9, $row_all, $product->getWeight())            //weight
						->setCellValueByColumnAndRow(10, $row_all, $product->getStatus())           //status
						->setCellValueByColumnAndRow(11, $row_all, $product->getVisibility())       //visibility
						->setCellValueByColumnAndRow(13, $row_all, $product->getDescription())      //description
						->setCellValueByColumnAndRow(14, $row_all, $product->getShortDescription()) //short_description
						->setCellValueByColumnAndRow(18, $row_all, $product_date);                  //date of creation

					/** Set categories */
					if (!empty($product->getCategoryids())){
						$this->setCategoryNames($objPHPExcel, $product->getCategoryIds(), $row_all);
					}
					/** Set stock fields */
					$this->setStockFields($objPHPExcel, $product, $row_all);
					/** Set attributes */
					$this->setSpecificAttributes($objPHPExcel, $product, $row_all);
				}
				/** Set children */
				if ($product->getTypeId() == 'configurable' && $page!='products'){
					$children = $product->getTypeInstance()->getUsedProducts($product->getId());
					if(!empty($children)){
						$row_assoc = $this->setChildrenProducts($objPHPExcel, $children, $product, $row_assoc, $assoc_index);
					}
				}
				$row_all += 1;
			}
		}
		/** Save Excel file */
		try{
			$objPHPExcel->setActiveSheetIndex(0);
			if($fileFormat == 'xlsx'){
				$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
				header('Content-type: application/vnd.ms-excel');
				header('Content-Disposition: attachment; filename="products.xlsx"');
				$objWriter->save('php://output');
			} else{
				$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'CSV');
				header('Content-type: text/csv');
				header('Content-Disposition: attachment; filename="products.csv"');
				$objWriter->save('php://output');
			}
		} catch(Exception $e){
			throw new Exception('Error while saving Excel file.');
		}	
	}
	private function checkIfInChosenCat($catIds, $chosenCat){
		if($chosenCat == 'all_categories'){
			return true;
		}
		foreach($catIds as $category){
			if($category == $chosenCat){
				return true;
			}
		}
		return false;
	}
	private function checkIfInDateRange($product_date, $date_start, $date_end){
		$product_date = new DateTime($product_date);
		$date_start = new DateTime($date_start);
		$date_end = new DateTime($date_end);
		if (($product_date >= $date_start && $product_date <= $date_end)||($product_date <= $date_start && $product_date >= $date_end)){
			return true;
		}
		return false;
	}

	private function setHeaderNames($objPHPExcel, $page, $assoc_index){
		if($page!='assoc'){
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
				->setCellValueByColumnAndRow(18, 1, "created_at");
			$objPHPExcel->setActiveSheetIndex(0)->setTitle("All products");
		}
		if($page=='both'){
			$objPHPExcel->createSheet($assoc_index);
		}
		if($page!='products'){
			$objPHPExcel->setActiveSheetIndex($assoc_index)
				->setCellValueByColumnAndRow(0, 1, "sku")
				->setCellValueByColumnAndRow(1, 1, "_super_products_sku")
				->setCellValueByColumnAndRow(2, 1, "_super_attribute_code")
				->setCellValueByColumnAndRow(3, 1, "_super_attribute_option");
			$objPHPExcel->setActiveSheetIndex($assoc_index)->setTitle("Association");
		}
	}

	private function setChildrenProducts($objPHPExcel, $children, $product, $row_assoc, $assoc_index){
		Mage::app()->setCurrentStore(Mage_Core_Model_App::ADMIN_STORE_ID);
		$productAttributeOptions = $product->getTypeInstance(true)->getConfigurableAttributesAsArray($product);
		$objPHPExcel->setActiveSheetIndex($assoc_index)
			->setCellValueByColumnAndRow(0, $row_assoc, $product->getSku());
		foreach($productAttributeOptions as $productAttribute){
			$atr_code = $productAttribute['attribute_code'];
			$objPHPExcel->setActiveSheetIndex($assoc_index)
				->setCellValueByColumnAndRow(0, $row_assoc, $product->getSku());
			$objPHPExcel->setActiveSheetIndex($assoc_index)
				->setCellValueByColumnAndRow(2, $row_assoc, $atr_code);
			$row_assoc+=1;
			foreach($children as $child){
				$child = Mage::getModel('catalog/product')->loadByAttribute('sku', $child['sku']);
				$objPHPExcel->setActiveSheetIndex($assoc_index)
					->setCellValueByColumnAndRow(1, $row_assoc, $child->getSku());
				$objPHPExcel->setActiveSheetIndex($assoc_index)
					->setCellValueByColumnAndRow(2, $row_assoc, $atr_code);
				$objPHPExcel->setActiveSheetIndex($assoc_index)
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
		"recurring_profile","custom_design","custom_design_from","custom_design_to","custom_layout_update","page_layout","options_container","gift_message_available", "manufacturer", "featured_product", "best_selling",
		"open_amount_min", "open_amount_max", "aht_deal_qty", "is_redeemable", "use_config_is_redeemable", "lifetime", "use_config_lifetime", "email_template", "use_config_email_template", "allow_message", "use_config_allow_message",
		"postmethods", "fit_size", "posting_days", "gift_wrapping_available", "gift_wrapping_price");

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
				->setCellValueByColumnAndRow(19+$index, 1, $attributeCode);
			if(!empty($product->getAttributeText($attributeCode))){
				$objPHPExcel->setActiveSheetIndex(0)
					->setCellValueByColumnAndRow(19+$index, $row_all, $product->getAttributeText($attributeCode));
			} elseif(!empty($product->getData($attributeCode))){
				$objPHPExcel->setActiveSheetIndex(0)
					->setCellValueByColumnAndRow(19+$index, $row_all, $product->getData($attributeCode));
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
