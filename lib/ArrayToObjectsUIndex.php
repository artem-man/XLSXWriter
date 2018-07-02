<?php
/*
	This class is designed for store unique objects from an array of parameters
	in template-way - objects are created only if necessary by callback function.

	@author Artem Myrhorodskyi
*/
class ArrayToObjectsUIndex
{
	private $hashIndexes = array();
	private $allowedKeys;
	private $counter = 0;

	public $objStorage = array();

/*
	Create a new ArrayToObjectsUIndex

	@param $allowedKeys - array of parameters applicable for creating an object.
	                      Usefull for build unique key from for redundant data.

	@param $startIdx - default value for internal counter
*/

	public function __construct(array $allowedKeys=array(), $startIdx=0)
	{
		$this->allowedKeys = array_flip($allowedKeys);
		$this->counter = $startIdx;
	}

/*
	This function tries to search a previously stored object from incoming
	array of parameters, otherwise it creates a new one.

	@param $data - array of parameters for creating an objects

	@param $func_create_object - callback function that should create an object
	                             from the incoming array of parameters

	@return index of unique object

	@usage:
		$objArrayToObjectsUIndex->lookup($incomingArray, function($filtered_data) {
			return new SomeObject($filtered_data);
		});

*/
	public function lookup(array $data, $func_create_object)
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

/*
	Return array of stored objects
*/

	public function data()
	{
		return $this->objStorage;
	}

/*
	Return internal counter. Attention! To get the actual number of saved objects
	please use count($this->data())
*/

	public function count()
	{
		return $this->counter;
	}
}
