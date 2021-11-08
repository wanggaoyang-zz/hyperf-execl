# hyperf-execl
*  安装
*  composer require hyperf-wgy/excel
*  使用示例
*  $data =[
*       ['field1' => 11111,'field2' => 22222],
*       ['field1' => 111111,'field2' => 222222],
*  ];
*  $field = ['标题1' => 'field1','标题2' => 'field2']
*  make(Excel::class)->export($data, $field, './runtime/excel/' . date('y-m-d'), '测试数据.xlsx');
