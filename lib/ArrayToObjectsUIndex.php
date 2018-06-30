<?php
/*
 * @Artem Myrhorodskyi
 * */
class ArrayToObjectsUIndex
{
	private $hashIndexes = array();
	private $allowedKeys;
	private $counter = 0;

	public $objStorage = array();

	public function __construct(array $allowedKeys=array(), $startIdx=0)
	{
		$this->allowedKeys = array_flip($allowedKeys);
		$this->counter = $startIdx;
	}

	public function lookup($data, $func_create_object)
	{
		if (empty($this->allowedKeys)) {
			$vals = $data;
		}
		else {
			$vals = array_intersect_key($data, $this->allowedKeys);
		}

		if (empty($vals)) {
			return false;
		}

		ksort($vals);
		$needle = json_encode($vals);
		if (array_key_exists($needle, $this->hashIndexes)) {
			$idx = $this->hashIndexes[$needle];
		}
		else {
			$idx = $this->counter;
			$this->counter++;

			$this->hashIndexes[$needle] = $idx;
			$this->objStorage[$idx] = call_user_func($func_create_object, $vals);
		}
		return $idx;
	}

	public function data()
	{
		return $this->objStorage;
	}

	public function count()
	{
		return $this->counter;
	}
}

