# phpexcel

*PHPExcel module for Kohana 3.0.x*

- **Module URL:** <http://github.com/dshovchko/kohana-phpexcel>
- **Compatible Kohana Version(s):** 3.0.x

## Description

Kohana framework helper class to make spreadsheet creation easier

## Instalation

Place the module in the modules/phpexcel.
    
Install library PHPExcel. In the application/vendor/ create a directory phpexcel. Download from official website (http://phpexcel.codeplex.com/) the latest version of the library PHPExcel, unzip and copy the entire contents of the directory classes/ in the application/vendor/.
    
In the application/bootstrap.php add module loading
    
    Kohana::modules(array(
    ...
        'phpexcel'   => MODPATH.'phpexcel',
    ));

## Usage

    $ws = new Spreadsheet(array(
	'author'       => 'Kohana-PHPExcel',
	'title'	       => 'Report',
	'subject'      => 'Subject',
	'description'  => 'Description',
    ));
    
    $ws->set_active_sheet(0);
    $as = $ws->get_active_sheet();
    $as->setTitle('Report');
    
    $as->getDefaultStyle()->getFont()->setSize(9);
    
    $as->getColumnDimension('A')->setWidth(7);
    $as->getColumnDimension('B')->setWidth(40);
    $as->getColumnDimension('C')->setWidth(12);
    $as->getColumnDimension('D')->setWidth(10);
    
    $sh = array(
	1 => array('Day','User','Count','Price'),
	2 => array(1, 'John', 5, 587),
	3 => array(2, 'Den', 3, 981),
	4 => array(3, 'Anny', 1, 214)
    );
    
    $ws->set_data($sh, false);
    $ws->send(array('name'=>'report', 'format'=>'Excel5'));
