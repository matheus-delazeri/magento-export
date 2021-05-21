<?php

class Matheus_Export_IndexController extends Mage_Adminhtml_Controller_Action
{
    public function indexAction()
    {
        $today = date('Y-m-d');
        $url = $this->getUrl('export_products/start/index');
        $urlValue = Mage::getSingleton('core/session')->getFormKey();
        $categories = $this->getCategories();

        $block_content = "
        <h4>Export file format: </h4>
        <form action='$url' method='post'>
	    <select name='file_type'>
                <option value='xlsx'>XLSX</option>
                <option value='csv'>CSV</option>
            </select>
	    <p>* CSV format only supports one page per time (or products or association)</p>
	    <h4>Select the date range for export:</h4>
            <label for='date_start'>Export products from: </label>
            <br>
            <input type='date' name='date_start' id='date_start' value='$today'>
            <br><br>
            <label for='date_end'>To: </label>
            <br>
            <input type='date' name='date_end' id='date_end' value='$today'>
            <br><br>
            <h4>Select which pages should be exported:</h4>
            <select name='sheets'>
                <option value='both'>All Pages</option>
                <option value='assoc'>Only Association</option>
                <option value='products'>Only Products</option>
            </select>
            <br><br>
            <h4>Select from which category the products should be exported:</h4>
            <select name='categories'>
                <option value='all_categories'>All Categories</option>";
            
        foreach($categories as $id=>$category){
            $block_content .= "<option value='$id'>$category</option>";
        }     
        $block_content .= "
            </select>
            <br><br>
            <input type='hidden' name='form_key' value='$urlValue'>
            <input type='submit' class='btn-export' id='submit' value='Export'>
        </form>
        
        <style type='text/css'>
        .btn-export{
            display: block;
            border: 0;
            width: 80px;
            background: #4E9CAF;
            padding: 5px 0%;
            text-align: center;
            border-radius: 5px;
            color: white;
            font-weight: bold;
            cursor: pointer;
            line-height: 25px; 
        }
        </style>"; 
        $this->loadLayout();

        $this->_setActiveMenu('catalog/matheus');
        $block = $this->getLayout()
            ->createBlock('core/text', 'export-block')
            ->setText($block_content);

        $this->_addContent($block);
        $this->renderLayout();
    }

    private function getCategories(){
        $category = Mage::getModel('catalog/category');
		$catTree = $category->getTreeModel()->load();
		$catIds = $catTree->getCollection()->getAllIds();
		if ($catIds){
            $catNames = array();
			foreach ($catIds as $id){
				$cat = Mage::getModel('catalog/category');
				$cat->load($id);
                $catNames[$id] = $cat->getName();
			} 
            return $catNames;
		} 
    }
}
