<?php

class Matheus_Export_IndexController extends Mage_Adminhtml_Controller_Action
{
    public function indexAction()
    {
        $url = $this->getUrl('export_products/start/index');
        $url_value = Mage::getSingleton('core/session')->getFormKey();
        $block_content = "
        <h4>Select the date range for export:</h4>
        <form action='$url' method='post'>
            <label for='date_start'>Export products from: </label>
            <br>
            <input type='date' name='date_start' id='date_start'>
            <br><br>
            <label for='date_end'>To: </label>
            <br>
            <input type='date' name='date_end' id='date_end'>
            <br><br>
            <input type='hidden' name='form_key' value='$url_value'>
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
}