<?php
/**
 * Class for create temporary file as alternative to tmpfile() function with autodelete
 *
 * @author  Artem Myrhorodskyi
 * @thanks Aleksandr Denisyuk <a@denisyuk.by>
 */
class TempFile
{
    /**
     * @var string $filename Full path to temporary file
     */
    private $filename = null;

    /**
     * Create instance a temporary file and register auto delete function
     *
     */
    public function __construct()
    {
        $this->filename = $this->create();

        register_shutdown_function([$this, 'delete']);
    }

    /**
     * Create file with unique name in temp directory
     *
     * @throws \Error
     * @return string
     */
    private function create()
    {
        $filename = tempnam(sys_get_temp_dir(), 'php');

        if (!$filename) {
            throw new \Error('The function tempnam() could not create a file in temporary directory.');
        }

        return $filename;
    }

    /**
     * Force delete a temp file
     *
     * @return bool
     */
    public function delete()
    {
    	if ($this->filename && file_exists($this->filename)) {
	        $ret = @unlink($this->filename);
	        $this->filename = null;
	        return $ret;
        }
        return true;
    }

    /**
     * Transform the object in the filename
     *
     * @return string
     */
    public function __toString()
    {
        return $this->filename;
    }
}

/**
 * This class return file path to temporary file instead of object,
 * usefull for configuration with strict types check
 *
 * @author  Artem Myrhorodskyi
 * @license MIT
 */

class TempFileCreator
{
	private static $tmpfileList = array();

	public static function tempFilename()
	{
		$tmpfile = new TempFile();
		self::$tmpfileList[ $tmpfile->__toString() ] = $tmpfile;
		return $tmpfile->__toString();
	}

	public static function delete($filename)
	{
		if (isset($tmpfileList[$filename])) {
			$tmpfileList[$filename]->delete();
			unset($tmpfileList[$filename]);
		}
	}
}
