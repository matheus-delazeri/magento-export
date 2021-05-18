<?php

class Matheus_Import_StartController extends Mage_Core_Controller_Front_Action {
	public function indexAction(){
		/** Set default timezone */
		date_default_timezone_set('America/Bahia');
		/** Set file directory */
		$target_dir = __DIR__.'/../sheets/';
		$this->fileName = $target_dir . basename($_FILES["fileToUpload"]["name"]);
		$excelFileType = strtolower(pathinfo($this->fileName,PATHINFO_EXTENSION));
		$uploadOk = 1;

		/** Check file size */
		if ($_FILES["fileToUpload"]["size"] > 500000) {
			echo date('H:i:s')," Your file is too large.<br>";
			$uploadOk = 0;
		}

		/** Check if $uploadOk is set to 0 by an error */
		if ($uploadOk == 0) {
			echo date('H:i:s')," Sorry, your file was not uploaded.<br>";
		/** if everything is ok, try to upload file */
		} else {
			if (move_uploaded_file($_FILES["fileToUpload"]["tmp_name"], $this->fileName)) {
				echo date('H:i:s')," The file ". htmlspecialchars( basename( $_FILES["fileToUpload"]["name"])). " has been uploaded.<br>";
				echo date('H:i:s')," Starting the import process...<br>";
				$this->setProductInfos();
			} else {
				echo date('H:i:s')," Sorry, there was an error uploading your file.<br>";
			}
		}

	}
	/** Set ALL infos again */
	private function setProductInfos(){
		require_once dirname(__FILE__).'/../Classes/PHPExcel.php';
		$inputFileType = PHPExcel_IOFactory::identify($this->fileName);
		$objReader = PHPExcel_IOFactory::createReader($inputFileType);
		$objPHPExcel = $objReader->load($this->fileName);
		$worksheet  = $objPHPExcel->getActiveSheet();
		$highestRow = $worksheet->getHighestRow();

		$childrenArray = array();
		/** Build the product */
		for ($i = 2; $i <= (int)$highestRow; $i++){
		    Mage::app()->setCurrentStore(Mage_Core_Model_App::ADMIN_STORE_ID);
		    $product = Mage::getModel('catalog/product'); 
	            $attributeSet = $product->getDefaultAttributeSetId();
		    /** Set values */
		    $product
			->setStoreId(1)
			->setWebsiteIds(array(1))
			->setTaxClassId(0)
			->setAttributeSetId($attributeSet)
			->setTypeId(($worksheet->getCellByColumnAndRow(3, $i)->getValue()))
			->setSku(($worksheet->getCellByColumnAndRow(7, $i)->getValue()))
			->setName(($worksheet->getCellByColumnAndRow(8, $i)->getValue()))
			->setPrice(($worksheet->getCellByColumnAndRow(10, $i)->getValue()))
			->setWeight(($worksheet->getCellByColumnAndRow(12, $i)->getValue()))
			->setStatus(($worksheet->getCellByColumnAndRow(14, $i)->getValue()))
			->setVisibility(($worksheet->getCellByColumnAndRow(15, $i)->getValue()))
			->setDescription(($worksheet->getCellByColumnAndRow(20, $i)->getValue()))
			->setShortDescription(($worksheet->getCellByColumnAndRow(21, $i)->getValue()))
			->setCategoryIds(array(2))
		        ->setStockData([
				'is_in_stock' => $worksheet->getCellByColumnAndRow(23, $i)->getValue(),
				'qty' => $worksheet->getCellByColumnAndRow(22, $i)->getValue()
			]);

		    /** Save product */
		    try{
			    $product->save();
		    } catch(Exception $e){
			    $newProduct = $product;
			    $product->delete();
			    $newProduct->save();
			}
		    if($product->getTypeId() == 'grouped'){
			    $children = $worksheet->getCellByColumnAndRow(4, $i)->getValue();
			    if($children !== NULL && !is_int($children) && !is_float($children)){
				    $childrenArray[$product->getId()] = $children;
			    }
		    } 
		}
		echo date('H:i:s')," Linking grouped/configurable subproducts...<br>";
		$this->getConfigAndGroupedProducts($childrenArray);
		echo date('H:i:s')," Products successfully imported!<br>";
	}
	private function getConfigAndGroupedProducts($childrenArray){
		foreach ($childrenArray as $groupedId => $children) {
			$childrenIds = [];
			$productsLinks = Mage::getModel('catalog/product_link_api');
			$childrenSku = explode('\/', $children);
			for($i = 0; $i < count($childrenSku); $i++){
				$childrenIds[$i] = Mage::getModel('catalog/product')->getIdBySku($childrenSku[$i]); 
			}
			foreach($childrenIds as $id) {
				$productsLinks->assign('grouped', $groupedId, $id);
			}
		}
	}
}
