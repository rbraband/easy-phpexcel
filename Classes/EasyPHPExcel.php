<?php

class EasyPHPExcel{

	protected $title;
	protected $description;
	protected $creator;

	private $sheets;
	private $header;
	private $rows;

	private $columnCount;
	private $currentSheet;

	private $objPHPExcel;
	private $objWorkSheet;

	private $columnNames;


	/**
	 * @param string $title
	 * @param string $description
	 * @param string $creator
	 */
	function __construct($title = '', $description = '', $creator = '')
	{

		$this->title = $title;
		$this->description = $description;
		$this->creator = $creator;

		$this->columnCount 	= array(0 => 0);
		$this->currentSheet = 0;

		$this->sheets   = array(0 => array('title' => $title));
		$this->header   = array(0 => array());
		$this->rows     = array(0 => array());

		$this->objPHPExcel = new PHPExcel();
		$this->objPHPExcel->getProperties()->setCreator($creator);
		$this->objPHPExcel->getProperties()->setLastModifiedBy($creator);
		$this->objPHPExcel->getProperties()->setTitle($title);
		$this->objPHPExcel->getProperties()->setDescription($description);

		$this->columnNames = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

	}


	/**
	 * @param array $header
	 *
	 * @return $this
	 */
	public function setHeader($header)
	{

		if (!isset($this->header[$this->currentSheet]))
			$this->header[$this->currentSheet] = array();
		
		$this->header[$this->currentSheet] = $header;
		
		if (!isset($this->columnCount[$this->currentSheet]))
			$this->columnCount[$this->currentSheet] = 0;
		
		if (count($header) > $this->columnCount[$this->currentSheet]) {

			$this->columnCount[$this->currentSheet] = count($header);

		}

		return $this;

	}
	
	public function setSheetTitle($title)
	{
		$this->sheets[$this->currentSheet]['title'] = $title;
		
		return $this;
	}
	
	public function setActiveSheet($activeSheet = 0)
	{
		if (isset($this->sheets[$activeSheet])) {
			$this->currentSheet = $activeSheet;
			
			return $this;
		}
	}
	
	public function addSheet($title)
	{
		$sheets = array_push($this->sheets, array('title' => $title));
		return $sheets -1;
	}


	/**
	 * @param $rows
	 *
	 * @return $this
	 */
	public function addRows($rows)
	{
		foreach($rows as $row) {

			$this->addRow($row);
		}

		return $this;
	}

	/**
	 * @param array $row
	 *
	 * @return $this
	 */
	public function addRow($row)
	{

		if (!isset($this->rows[$this->currentSheet]))
			$this->rows[$this->currentSheet] = array();
		
		$this->rows[$this->currentSheet][] = $row;
		
		if (!isset($this->columnCount[$this->currentSheet]))
			$this->columnCount[$this->currentSheet] = 0;

		if (count($row) > $this->columnCount[$this->currentSheet])
			$this->columnCount[$this->currentSheet] = count($row);

		return $this;

	}

	/**
	 * @param int $index
	 *
	 * @return string
	 */
	private function getColumnCharacter($index)
	{

		return substr($this->columnNames, $index, 1);

	}


	/**
	 * Generate based on rows and header the document
	 *
	 * @return $this
	 */
	private function buildDocument()
	{

		// sheets		
		
		foreach($this->sheets as $currentSheet => $sheet) {
			$currentRow = 1;
			
			// header		
		
			if ($currentSheet !== 0)
				$this->objWorkSheet = $this->objPHPExcel->createSheet($currentSheet);
			else
				$this->objWorkSheet = $this->objPHPExcel->getActiveSheet();
			
			$this->objWorkSheet->setTitle($sheet['title']);

			if (count($this->header[$currentSheet]) > 0) {

				for($i = 0; $i < $this->columnCount[$currentSheet]; $i++) {
					$currentCell = $this->getColumnCharacter($i) . $currentRow;
					$currentData = '';
					if (isset($this->header[$currentSheet][$i])) {
						$currentData = $this->header[$currentSheet][$i];
					}
					$this->objWorkSheet->SetCellValue($currentCell , $currentData);
					$this->objWorkSheet->getStyle($currentCell)->getFont()->setBold(true);
				}
				$currentRow++;
			}
		
			// rows		
		
			if (count($this->rows[$currentSheet]) > 0) {
				foreach($this->rows[$currentSheet] as $row) {
					for($i = 0; $i < $this->columnCount[$currentSheet]; $i++) {
						$currentCell = $this->getColumnCharacter($i) . $currentRow;
						$currentData = '';
						if (isset($row[$i])) {
							$currentData = $row[$i];
						}
						$this->objWorkSheet->SetCellValue($currentCell , $currentData);
					}
					$currentRow++;
				}
			}

			$this->applyAutoSizing();
		}
		
		return $this;
	}


	/**
	 * Save the document to file system
	 *
	 * @param $file
	 *
	 * @throws PHPExcel_Writer_Exception
	 */
	public function save($file)
	{
		$this->buildDocument();
		$this->objPHPExcel->setActiveSheetIndex(0);
		$objWriter = new PHPExcel_Writer_Excel2007($this->objPHPExcel);
		$objWriter->save($file);
	}

	/**
	 * Apply auto sizing
	 */
	private function applyAutoSizing()
	{
		foreach (range('A', $this->objWorkSheet->getHighestDataColumn()) as $col) {

			$this->objWorkSheet->getColumnDimension($col)
							   ->setAutoSize(true);
		}

	}
}
