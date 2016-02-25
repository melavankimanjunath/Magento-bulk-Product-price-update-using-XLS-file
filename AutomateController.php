<?php

class Npm_Automate_AutomateController extends Mage_Adminhtml_Controller_Action
{
    /**
    * IndexAction load the form.
    * Give interface to upload xls file.
    **/

    public function indexAction()
    {
        $this->loadLayout()
            ->_addContent(
            $this->getLayout()
            ->createBlock('npm_automate/upload')
            ->setTemplate('automate/form.phtml'))
            ->renderLayout();
    }
    /**
    * Post action loads when form submit clicks.
    * Verify the xls file and call the @getFileData and @callToPriceUpdater functions
    **/

    public function postAction()
    {
        $post = $this->getRequest()->getPost();
        $fileType = '';
        $fullname = '';
        if (isset($_FILES['file']['name']) && $_FILES['file']['name'] != '') {
            try
            {
                $uploader = new Varien_File_Uploader('file');
                $uploader->setAllowedExtensions(array('xls','xlsx'));
                $uploader->setAllowCreateFolders(true);
                $uploader->setAllowRenameFiles(false);
                $uploader->setFilesDispersion(false);
                $path = Mage::getBaseDir('media').DS.'CSV'.DS;
                $fname = $_FILES['file']['name'];
                $fullname = $path.$fname;
                $uploader->save($path, $fname);
            }
            catch (Exception $e)
            {
                $fileType = "Invalid file format";
            }
        } else{
            die('invalid file 1');
        }
        if ($fileType == "Invalid file format") {
            Mage::getSingleton('adminhtml/session')->addError("Invalid file format");
            $this->_redirect('*/*/');
            unlink($fullname);
            return;
        }
        if ($fullname != "") {
            /**
            * Extract Data from File and get it in @fileDataRecords.
            **/
            $fileDataRecords = $this->getFileData($fullname);
            /**
            * Method, update the product price.
            **/
            $result = $this->callToPriceUpdater($fileDataRecords);
        }
        Mage::getSingleton('adminhtml/session')->addSuccess('Product Price are Updated Successfully');
        Mage::getSingleton('adminhtml/session')->setFormData(false);
        $this->_redirect('*/*');
    }

    public function callToPriceUpdater($fileDataRecords)
    {
        $products = Mage::getModel('catalog/product');
        $count = 0;
        foreach($fileDataRecords as $_data){
            $website = Mage::getModel('core/website')->load($_data[0][3])->getStoreIds();
            $prod = $products->loadByAttribute('sku',$_data[0][0]);
            $count = $count + 1;
            if ($prod) {
                foreach ($website as  $store) {
                    $prod->setStoreId($store)->setPrice($_data[0][1])->save();
                }
            }
         }
         return $count;
    }

    public function convertWebsiteNameById($websiteName)
    {
        $websiteId = null;
        foreach (Mage::app()->getWebsites() as  $websites) {
            if ($websites->getName() == $websiteName) {
                $websiteId = $websites->getId();
            }
        }
        return $websiteId;
    }

    public function getFileData($fullname)
    {
        require( Mage::getBaseDir('lib').'/phpexcel/PHPExcel.php');
        require(Mage::getBaseDir('lib').'/phpexcel/PHPExcel/IOFactory.php');
        try {
                $inputFileType = PHPExcel_IOFactory::identify($fullname);
                $objReader = PHPExcel_IOFactory::createReader($inputFileType);
                $objPHPExcel = $objReader->load($fullname);
            } catch(Exception $e) {
                die('Error loading file "'.pathinfo($fullname,PATHINFO_BASENAME).'": '.$e->getMessage());
            }
        $sheet = $objPHPExcel->getSheet(0);
        $highestRow = $sheet->getHighestRow();
        $highestColumn = $sheet->getHighestColumn();
        $fileDataRecords =array();

        for ($row = 1; $row <= $highestRow; $row++)
        {
            //  Read a row of data into an array
            $rowData = $sheet->rangeToArray(
                'A' . $row . ':' . $highestColumn . $row, NULL, TRUE, FALSE
                );

            if ($rowData[0][0]!= null && $rowData[0][1]!=null && $rowData[0][2]!=null) {
                $websiteId = $this->convertWebsiteNameById($rowData[0][2]);
                if ($websiteId == null) {
                    $fileDataRecords = array();
                    Mage::getSingleton('adminhtml/session')->addError('Enter valid website name and try again later (Valid are AUD, CAD, USD)');
                    Mage::getSingleton('adminhtml/session')->setFormData(false);
                    return Mage::helper('adminhtml')->getUrl('adminhtml/automate/index');
                    break;
                }
                $rowData[0][3] = $websiteId;
                array_push($fileDataRecords, $rowData);
            }
        }
        return $fileDataRecords;
    }
}