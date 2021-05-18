<?php
class Matheus_Import_IndexController extends Mage_Core_Controller_Front_Action {
	public function indexAction() {
		echo "
<form action='import/start' method='post' enctype='multipart/form-data'>
  Select file to upload:
  <input type='file' name='fileToUpload' id='fileToUpload'>
  <input type='submit' value='Upload file' name='submit'>
</form>";
	}
}



