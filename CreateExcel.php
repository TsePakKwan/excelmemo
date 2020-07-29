<?php
namespace App\Model\Mst;

use Illuminate\Support\Facades\Storage;

//PhpOffice
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Style\Style;

use App\Model\MstProc;

use Log;


/**
 * Excel操作
 * テンプレートからデータを入れる、出力フォルダに出力する
 */
class CreateExcel {
    /**
     * @var Spreadsheet $_spreadsheet テンプレートからEXCEL入れ
     */
    protected $_spreadsheet;        
    /**
     * @var string $_templateDirectoryPathName テンプレートpath
     */
    protected $_templateDirectoryPathName;
    /**
     * @var string $_outputDirectoryPathName 出力のフォルダ
     */
    protected $_outputDirectoryPathName;        
    /**
     * @var string $title タイトル
     */
    protected $title;
    /**
     * @var string $_output_file_name 出力ファイル名    
     */
    protected $_output_file_name;
    /**
     * @var string $is_exists 出力ファイルの既存
     */
    protected $is_exists = false;
    /**
     * @var \Illuminate\Filesystem\FilesystemAdapter $disk
     */
    protected $disk;
    /**
     * @var string $full_path 実際に登録したpath
     */
    private $full_path;

    public function __construct($templateDirPathName, $outputDirPathName, $output_file_name) 
    {
        set_time_limit(0);
        $this->disk = Storage::disk('local');
        //実際に登録したpath
        $this->full_path = $this->disk->getAdapter()->getPathPrefix();

        //出力ファイル
        $url = $outputDirPathName . $output_file_name . '.xlsx';
        $this->is_exists = $this->disk->exists($url);
        //出力ファイル名
        $this->_output_file_name = $output_file_name;
        //テンプレートpath
        $this->_templateDirectoryPathName = $templateDirPathName;
        //出力のフォルダ
        $this->_outputDirectoryPathName = $outputDirPathName;

        $this->_spreadsheet = IOFactory::load($this->full_path . $templateDirPathName);
        
    }

    /**
     * -------------------------------------------------------------
     * クラス設定/取得関連
     * -------------------------------------------------------------
     */

    /**
     * excelのタイトルを設定
     * @param string $title
     */
    public function setTitle(string $title) 
    {
        $this->title = $title;
    }

    /**
     * excelのタイトルを取得
     * @return string
     */
    public function getTitle() 
    {
        return $this->title;
    }

    /**
     * 出力ファイルの名前を設定
     * @param string $output_file_name
     */
    public function setOutputFileName(string $output_file_name) 
    {
        $this->_output_file_name = $output_file_name;
    }

    /**
     * 出力ファイルの名前を取得
     * @return string
     */
    public function getOutputFileName() 
    {
        return $this->_output_file_name;
    }

    /**
     * -------------------------------------------------------------
     * Excel　シート操作関連
     * -------------------------------------------------------------
     */


    public function getSpreadSheetIndexByName($name) 
    {
        return $this->_spreadsheet->getIndex(
                        $this->_spreadsheet->getSheetByName($name)
        );
    }

    /**
     * PhpSpreadsheetのクラスStyleをapplyFromArray用の配列に変更
     * @param Style $style
     * @return array
     */
    protected function getFormArray(Style $style) 
    {
        return [
            //フォント
            'font' => [
                //フォント名
                'name' => $style->getFont()->getName(),
                //文字サイズ
                'size' => $style->getFont()->getSize(),
                //ボールド体
                'bold' => $style->getFont()->getBold(),
                //イタリック体
                'italic' => $style->getFont()->getItalic(),
                //下線
                'underline' => $style->getFont()->getUnderline(),
                //取り消し線
                'strikethrough' => $style->getFont()->getStrikethrough(),
                'color' => [
                    //文字色
                    'rgb' => $style->getFont()->getColor()->getRGB()
                ]
            ],
            //罫線
            'borders' => [
                'bottom' => [
                    'borderStyle' => $style->getBorders()->getBottom()->getBorderStyle(),
                    'color' => [
                        'rgb' => $style->getBorders()->getBottom()->getColor()->getRGB()
                    ]
                ],
                'top' => [
                    'borderStyle' => $style->getBorders()->getTop()->getBorderStyle(),
                    'color' => [
                        'rgb' => $style->getBorders()->getTop()->getColor()->getRGB()
                    ]
                ],
                'left' => [
                    'borderStyle' => $style->getBorders()->getLeft()->getBorderStyle(),
                    'color' => [
                        'rgb' => $style->getBorders()->getLeft()->getColor()->getRGB()
                    ]
                ],
                'right' => [
                    'borderStyle' => $style->getBorders()->getRight()->getBorderStyle(),
                    'color' => [
                        'rgb' => $style->getBorders()->getRight()->getColor()->getRGB()
                    ]
                ],
                'diagonal' => [
                    'borderStyle' => $style->getBorders()->getDiagonal()->getBorderStyle(),
                    'color' => [
                        'rgb' => $style->getBorders()->getDiagonal()->getColor()->getRGB()
                    ]
                ],
                'diagonaldirection' => $style->getBorders()->getDiagonalDirection(),
            ],
            //配置
            'alignment' => [
                'horizontal' => $style->getAlignment()->getHorizontal(),        //水平                
                'vertical' => $style->getAlignment()->getVertical(),            //垂直
                'textRotation' => $style->getAlignment()->getTextRotation(),    //回転
                'wrapText' => $style->getAlignment()->getWrapText(),            //折返し
                'shrinkToFit' => $style->getAlignment()->getShrinkToFit(),        //縮小して全体を表示
            ],
            //
            'numberFormat' => [
                'formatCode' => $style->getNumberFormat()->getFormatCode(),
            ],
            'fill' => [
                'fillType' => $style->getFill()->getFillType(),
                'startColor' => [
                    'argb' => $style->getFill()->getStartColor()->getARGB(),
                ],
                'endColor' => [
                    'argb' => $style->getFill()->getEndColor()->getARGB(),
                ],
            ],
            'quotePrefix' => $style->getQuotePrefix()
        ];
    }

    /**
     * 指定した範囲($srcRange)のセルを配列として取得
     * 
     * @param Worksheet $sheet 操作するシート
     * @param string $srcRange コピー範囲、例: A27:B100
     * @return array
     */
    protected function rangeToCellArray(Worksheet $sheet, string $srcRange) 
    {
        // コピー範囲を検証します。
        if (!preg_match('/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/', $srcRange, $srcRangeMatch)) {
            return;
        }

        $srcColumnStart = $srcRangeMatch[1];
        $srcRowStart = $srcRangeMatch[2];
        $srcColumnEnd = $srcRangeMatch[3];
        $srcRowEnd = $srcRangeMatch[4];

        //ループのために、代わりにインデックスを変換する必要があります
        //注：このメソッドのように列は0ベースで1ベースではないため、1を減算する必要があります。

        $srcColumnStart = Coordinate::columnIndexFromString($srcColumnStart) - 1;
        $srcColumnEnd = Coordinate::columnIndexFromString($srcColumnEnd) - 1;
        
        $returnValue = [];

        for ($row = $srcRowStart; $row <= $srcRowEnd; $row++) {
            for ($col = $srcColumnStart; $col <= $srcColumnEnd; $col++) {
                //コピー
                $cell = $sheet->getCellByColumnAndRow($col, $row);
                $returnValue[$row][$col] = [
                    'value' => $cell->getValue(),
                    'style' => $this->getFormArray($cell->getStyle()),
                    'col_width' => $sheet->getColumnDimensionByColumn($col)->getWidth(),
                    'row_height' => $sheet->getRowDimension($row)->getRowHeight(),
                ];
            }
        }
        return $returnValue;
    }

    /**
     * 指定したセル（$dstCell）をはじめに$cellArrayの中身をペーストする
     * 
     * @param Worksheet $sheet 操作するシート
     * @param array $cellArray セルの配列
     * @param string $dstCell ペーストの最初セル、例: C27
     * @param array $option 値もペースト
     * @return void
     */
    protected function pasteCellArray(Worksheet $sheet, array $cellArray,string $dstCell, array $option = []) 
    {
        // ペースト先のセルを検証します。
        if (!preg_match('/^([A-Z]+)(\d+)$/', $dstCell, $destCellMatch)) {
            return;
        }
        if( isset($option['withValue']) === false ){
            $option['withValue'] = true;
        }

        if( isset($option['withStyle']) === false ){
            $option['withStyle'] = true;
        }
        
        $destColumnStart = Coordinate::columnIndexFromString($destCellMatch[1]);
        $destRowStart = $destCellMatch[2];

        $rowCount = 0;
        foreach ($cellArray as $row) {
            $rowIndex = $destRowStart + $rowCount;
            $colCount = 0;
            foreach ($row as $cell) {
                $colIndex = $destColumnStart + $colCount - 1;
                //ペースト
                $dstCell = Coordinate::stringFromColumnIndex($colIndex) . (string) ($rowIndex);
                if( $option['withValue'] === true ){
                    $sheet->setCellValue($dstCell, $cell['value']);
                }
                if( $option['withStyle'] === true ){
                    $sheet->getStyle($dstCell)->applyFromArray($cell['style']);
                }
                
                // 列の幅を設定しますが、列ごとに一回のみ
                if( $rowCount == 0 ){
                    $sheet->getColumnDimensionByColumn($colIndex)->setAutoSize(false);
                    $sheet->getColumnDimensionByColumn($colIndex)->setWidth($cell['col_width']);
                }
                
                //行の高さを設定、行ことに一回のみ
                if( $colCount == 0 ){
                    $sheet->getRowDimension($rowIndex)->setRowHeight($cell['row_height']);
                }

                $colCount++;
            }

            $rowCount++;
        }
        
    }

    /**
     * 指定した範囲($srcRange)をスタイルを含で指定したコピー、（$dstCell）にペーストする
     * 
     * @param Worksheet $sheet 操作するシート
     * @param string $srcRange コピー範囲、例: A27:B100
     * @param string $dstCell ペーストの最初セル、例: C27
     * @return void
     */
    protected function copyRange(Worksheet $sheet, string $srcRange, string $dstCell) 
    {
        //コピー
        $template = $this->rangeToCellArray($sheet, $srcRange);
        //ペースト
        $this->pasteCellArray($sheet, $template, $dstCell);
    }

    /**
     * -------------------------------------------------------------
     * Excel　出力操作関連
     * -------------------------------------------------------------
     */

    /**
     * 入力終えたテンプレートからEXCELを出力のフォルダに生成
     * @return String 生成しだEXCELの位置を戻します。
     */
    public function _excelSave() 
    {
        $writer = new Xlsx($this->_spreadsheet);

        $path = $this->_outputDirectoryPathName . $this->_output_file_name . '.xlsx';
        if ($path == $this->_templateDirectoryPathName && $this->is_exists === true) {
            $new_path = $this->_outputDirectoryPathName . $this->_output_file_name . 's.xlsx';
            $writer->save($this->full_path . $new_path);
            $this->disk->delete($path);
            $this->disk->rename($new_path, $path);
        } else {
            $writer->save($this->full_path . $path);
        }

        return $this->full_path . $path;
    }

    /**
     * 入力終えたテンプレートからEXCELを出力のフォルダに生成、
     * それをPDFに変換生成します
     * @return String 生成しだPDFの位置を戻します。
     */
    public function _pdfSave() 
    {
        //excel作成
        $excelfilepath = $this->createXls();
        // LibreOfficeでPDF化する
        $storagePath = $this->full_path . $this->_outputDirectoryPathName;

        $command = "export HOME=/tmp; "
            . "/bin/libreoffice "
            . "--headless "
            . "--nologo "
            . "--nofirststartwizard "
            . "--convert-to pdf:writer_pdf_Export "
            . "--outdir {$storagePath} "
            . $excelfilepath;
        exec($command);
        return $storagePath . $this->_output_file_name . '.pdf';
    }

    /**
     * 入力終えたテンプレートからEXCELを出力のフォルダに生成、
     * それをPNGに変換生成します
     * @return String 生成しだPNGの位置を戻します。
     */
    public function _pngSave() 
    {
        //pdf作成
        $pdffilepath = $this->createPdf();
        //png化する
        $storagePath = $this->full_path . $this->_outputDirectoryPathName;

        return MstProc::pdfToPng($pdffilepath, $storagePath, $this->_output_file_name);
    }

    /**
     * 上書き用
     * テンプレートExcelファイルから出力用のExcelファイルを作成します
     */
    public function createXls() 
    {
        //Excel 保存
        return $this->_excelSave();
    }

    /**
     * 上書き用
     * テンプレートExcelファイルから出力用のExcelファイルを作成しPDFに変換します
     */
    public function createPdf() 
    {
        //PDF保存
        return $this->_pdfSave();
    }

    /**
     * 上書き用
     * テンプレートExcelファイルから出力用のExcelファイルを作成しPNGに変換します
     */
    public function createPng() 
    {
        //PNG保存
        return $this->_pngSave();
    }

    /**
     * テンプレートからエクセルを作成時にエラーcatchするメソッド
     * 
     * @return string|null 成功の場合はエクセル作成後のリンク、失敗の場合はnull
     */
    public function tryCreateXls() 
    {
        // EXCEL処理エラーが出たらエラーVIEWに飛ぶ
        try {
            return $this->createXls();
        } catch (\Exception $ex) {
            Log::debug( $ex);
            return null;
        }
    }

    /**
     * エクセルをテンプレートからPDF作成時にエラーcatchするメソッド
     * 
     * @return String  成功の場合はpdf作成後のリンク、失敗の場合はnull
     */
    public function tryCreatePdf() 
    {
        // EXCEL処理エラーが出たらエラーVIEWに飛ぶ
        try {
            return $this->createPdf();
        } catch (\Exception $ex) {
            Log::debug( $ex);
            return null;
        }
    }

    /**
     * エクセルをテンプレートからPNG作成時にエラーcatchするメソッド
     * 
     * @return String  成功の場合はpdf作成後のリンク、失敗の場合はnull
     */
    public function tryCreatePng() 
    {
        // EXCEL処理エラーが出たらエラーVIEWに飛ぶ
        try {
            return $this->createPng();
        } catch (\Exception $ex) {
            Log::debug( $ex);
            return null;
        }
    }

}
