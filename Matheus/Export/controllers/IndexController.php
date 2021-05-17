<?php

class Matheus_Export_IndexController extends Mage_Adminhtml_Controller_Action
{
    public function indexAction()
    {
        $url = $this->getUrl('export_products/start/index');
        $this->loadLayout();

        $this->_setActiveMenu('catalog/matheus');
        $block = $this->getLayout()
        ->createBlock('core/text', 'export-block')
        ->setText("<h4>Click in the button below to export the products to a .xlsx sheet</h4>
    <a style='
    display: block;
    width: 115px;
    height: 25px;
    background: #4E9CAF;
    padding: 10px;
    text-align: center;
    border-radius: 5px;
    color: white;
    font-weight: bold;
    line-height: 25px; text-decoration: none;' href='{$url}'>Export</a>");

        $this->_addContent($block);
        $this->renderLayout();
    }
}