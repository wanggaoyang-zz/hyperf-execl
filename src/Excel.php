<?php

namespace Wgy\Excel;

use Hyperf\HttpMessage\Stream\SwooleFileStream;
use Hyperf\HttpServer\Response;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Excel
{

    /**
     *  使用示例
     *  $data =[
     *       ['field1' => 11111,'field2' => 22222],
     *       ['field1' => 111111,'field2' => 222222],
     *  ];
     *  $field = ['标题1' => 'field1','标题2' => 'field2']
     *  make(Excel::class)->export($data, $field, './runtime/excel/' . date('y-m-d'), '测试数据.xlsx');
     *
     *
     * @param array $data
     * @param array $field //$field = ['标题1' => '数据字段1','标题2' => '数据字段2']
     * @param string $save_path
     * @param string $file_name
     * @param array $properties
     * @return \Psr\Http\Message\ResponseInterface
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @author wgy
     */
    public function export(array $data, array $field, string $save_path, string $file_name, array $properties = [])
    {
        $objPHPExcel = new Spreadsheet();
        $objPHPExcel = $this->setProperties($objPHPExcel, $properties);
        $objActSheet = $objPHPExcel->getActiveSheet();
        $field_name = array_keys($field);
        $field_column = array_values($field);
        //设置header
        $i = 0;
        foreach ($field_name as $value) {
            $cellName = self::stringFromColumnIndex($i) . '1';
            $objActSheet->setCellValue($cellName, $value)->calculateColumnWidths();
            $objActSheet->getColumnDimension(self::stringFromColumnIndex($i))->setWidth(15);
            ++$i;
        }
        //设置value
        $len = count($field_column);
        if (!empty($data)) {
            foreach ($data as $key => $item) {
                $row = 2 + ($key * 1);
                for ($i = 0; $i < $len; ++$i) {
                    $objActSheet->setCellValueExplicit(self::stringFromColumnIndex($i) . $row, $item[$field_column[$i]] ?? '', DataType::TYPE_STRING);
                }
            }
        }
        $objWriter = new Xlsx($objPHPExcel);
        $dir = iconv("UTF-8", "GBK", $save_path);
        if (!file_exists($dir)) {
            mkdir($dir, 0777, true);
        }
        $objWriter->save($save_path . DIRECTORY_SEPARATOR . $file_name);
        $objWriter->setPreCalculateFormulas(false);
        $objActSheet->disconnectCells();
        unset($objActSheet, $objWriter);
        $response = new Response();
        $response = $response->withHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        $response = $response->withHeader('Expires', '0');
        $response = $response->withHeader('Pragma', 'public');
        $response = $response->withHeader('Content-Transfer-Encoding', 'binary');
        $response = $response->withHeader('Content-Disposition', 'attachment;filename = ' . $file_name);
        $response = $response->withBody(new SwooleFileStream($save_path . DIRECTORY_SEPARATOR . $file_name));
        return $response;

    }

    /**
     * @param int $pColumnIndex
     * @return mixed|string
     * @author wgy
     */
    public static function stringFromColumnIndex($pColumnIndex = 0)
    {
        static $_indexCache = [];
        if (!isset($_indexCache[$pColumnIndex])) {
            if ($pColumnIndex < 26) {
                $_indexCache[$pColumnIndex] = chr(65 + $pColumnIndex);
            } elseif ($pColumnIndex < 702) {
                $_indexCache[$pColumnIndex] = chr(64 + ($pColumnIndex / 26)) . chr(65 + $pColumnIndex % 26);
            } else {
                $_indexCache[$pColumnIndex] = chr(64 + (($pColumnIndex - 26) / 676)) . chr(65 + ((($pColumnIndex - 26) % 676) / 26)) . chr(65 + $pColumnIndex % 26);
            }
        }
        return $_indexCache[$pColumnIndex];
    }

    /**
     * @param \PhpOffice\PhpSpreadsheet\Spreadsheet $objPHPExcel
     * @param array $properties
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet
     * @author wgy
     */
    public function setProperties(Spreadsheet $objPHPExcel, array $properties)
    {
        // 设置excel的属性：
        if (!empty($properties)) {
            if (!empty($properties['creator'])) {
                //创建人
                $objPHPExcel->getProperties()->setCreator($properties['creator']);
            }

            if (!empty($properties['last_modified'])) {
                //最后修改人
                $objPHPExcel->getProperties()->setLastModifiedBy($properties['last_modified']);
            }

            if (!empty($properties['title'])) {
                //标题
                $objPHPExcel->getProperties()->setTitle($properties['title']);
            }

            if (!empty($properties['subject'])) {
                //题目
                $objPHPExcel->getProperties()->setSubject($properties['subject']);
            }

            if (!empty($properties['description'])) {
                //描述
                $objPHPExcel->getProperties()->setDescription($properties['description']);
            }

            if (!empty($properties['keywords'])) {
                //关键字
                $objPHPExcel->getProperties()->setKeywords($properties['keywords']);
            }

            if (!empty($properties['category'])) {
                //种类
                $objPHPExcel->getProperties()->setCategory($properties['category']);
            }
        }
        return $objPHPExcel;
    }

}