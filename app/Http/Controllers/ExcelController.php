<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Excel;
use Session;
use Response;
use Storage;

ini_set('memory_limit', -1);
ini_set('max_input_time', -1);
ini_set('max_execution_time', -1);

class ExcelController extends Controller
{
    public function getViewPage()
    {
        return view('excel');
    }

    public function getDownloadPage()
    {
        return view('downloadpage');
    }

    public function upload(Request $request)
    {
        $filepath = $request->file('file')->getRealPath();
        $filename = explode(".", $request->file('file')->getClientOriginalName())[0];

        $rupper = $request->input('rupper');
        $rlower = $request->input('rlower') * (-1);
        $rstd = $request->input('rstandard');
        $gupper = $request->input('gupper');
        $glower = $request->input('glower') * (-1);
        $gstd = $request->input('gstandard');
        $bupper = $request->input('bupper');
        $blower = $request->input('blower') * (-1);
        $bstd = $request->input('bstandard');
        $excel = Excel::load($filepath)->get();
        $odata = $excel->toArray();
        $title = $excel->getTitle();
        $result1 = array();
        $result2 = array();
        $rmvarr = array();//red measure value
        $gmvarr = array();//green measure value
        $bmvarr = array();//blue measure value
        $chip = $request->input('chip');
        $pixel = $request->input('pixel');
        foreach ($odata as $v) {
            if ($v['color'] == "R") {
                array_push($rmvarr, str_replace(" V", "", $v['measure_value']));
            } elseif ($v['color'] == "G") {
                array_push($gmvarr, str_replace(" V", "", $v['measure_value']));
            } elseif ($v['color'] == "B") {
                array_push($bmvarr, str_replace(" V", "", $v['measure_value']));
            }
        }
        if (!$rstd) {
            $rstd = $this->median($rmvarr);
        }
        if (!$gstd) {
            $gstd = $this->median($gmvarr);
        }
        if (!$bstd) {
            $bstd = $this->median($bmvarr);
        }
        $rmin = $rstd * (1 + $rlower / 100);
        $rmax = $rstd * (1 + $rupper / 100);
        $gmin = $gstd * (1 + $glower / 100);
        $gmax = $gstd * (1 + $gupper / 100);
        $bmin = $bstd * (1 + $blower / 100);
        $bmax = $bstd * (1 + $bupper / 100);

        array_push($result1, ['Group', 'Rows', 'Lines', 'Color', 'Standard Value', 'Hi Limit', 'Low Limit', 'Measure Value', 'Result']);
        foreach ($odata as $v) {
            if ($v['color'] == "R") {
                if (str_replace(" V", "", $v['measure_value']) <= $rmax && str_replace(" V", "", $v['measure_value']) >= $rmin) {//pass
                    $array = array(
                        "group" => $v['group'],
                        "rows" => $v['rows'],
                        "lines" => $v['lines'],
                        "color" => $v['color'],
                        "standard_value" => $rstd,
                        "hi_limit" => $rmax,
                        "low_limit" => $rmin,
                        "measure_value" => str_replace(" V", "", $v['measure_value']),
                        "result" => "PASS"
                    );
                } else {//ng
                    $array = array(
                        "group" => $v['group'],
                        "rows" => $v['rows'],
                        "lines" => $v['lines'],
                        "color" => $v['color'],
                        "standard_value" => $rstd,
                        "hi_limit" => $rmax,
                        "low_limit" => $rmin,
                        "measure_value" => str_replace(" V", "", $v['measure_value']),
                        "result" => "NG"
                    );
                    $ngplace = array(
                        "group" => $v['group'],
                        "rows" => $v['rows'],
                        "lines" => $v['lines'],
                        "color" => $v['color']
                    );
                    array_push($result2, $ngplace);
                }
                array_push($result1, $array);
            } elseif ($v['color'] == "G") {
                if (str_replace(" V", "", $v['measure_value']) <= $gmax && str_replace(" V", "", $v['measure_value']) >= $gmin) {//pass
                    $array = array(
                        "group" => $v['group'],
                        "rows" => $v['rows'],
                        "lines" => $v['lines'],
                        "color" => $v['color'],
                        "standard_value" => $gstd,
                        "hi_limit" => $gmax,
                        "low_limit" => $gmin,
                        "measure_value" => str_replace(" V", "", $v['measure_value']),
                        "result" => "PASS"
                    );
                } else {//ng
                    $array = array(
                        "group" => $v['group'],
                        "rows" => $v['rows'],
                        "lines" => $v['lines'],
                        "color" => $v['color'],
                        "standard_value" => $gstd,
                        "hi_limit" => $gmax,
                        "low_limit" => $gmin,
                        "measure_value" => str_replace(" V", "", $v['measure_value']),
                        "result" => "NG"
                    );
                    $ngplace = array(
                        "group" => $v['group'],
                        "rows" => $v['rows'],
                        "lines" => $v['lines'],
                        "color" => $v['color']
                    );
                    array_push($result2, $ngplace);
                }
                array_push($result1, $array);
            } elseif ($v['color'] == "B") {
                if (str_replace(" V", "", $v['measure_value']) <= $bmax && str_replace(" V", "", $v['measure_value']) >= $bmin) {//pass
                    $array = array(
                        "group" => $v['group'],
                        "rows" => $v['rows'],
                        "lines" => $v['lines'],
                        "color" => $v['color'],
                        "standard_value" => $bstd,
                        "hi_limit" => $bmax,
                        "low_limit" => $bmin,
                        "measure_value" => str_replace(" V", "", $v['measure_value']),
                        "result" => "PASS"
                    );
                } else {//ng
                    $array = array(
                        "group" => $v['group'],
                        "rows" => $v['rows'],
                        "lines" => $v['lines'],
                        "color" => $v['color'],
                        "standard_value" => $bstd,
                        "hi_limit" => $bmax,
                        "low_limit" => $bmin,
                        "measure_value" => str_replace(" V", "", $v['measure_value']),
                        "result" => "NG"
                    );
                    $ngplace = array(
                        "group" => $v['group'],
                        "rows" => $v['rows'],
                        "lines" => $v['lines'],
                        "color" => $v['color']
                    );
                    array_push($result2, $ngplace);
                }
                array_push($result1, $array);
            }
        }
        Session::put('excel1', $result1);
        Session::put('excel2', $result2);
        Session::put('chip', $chip);
        Session::put('pixel', $pixel);
        Session::put('fname', $filename);
        Session::put('title', $title);
        return $this->getDownloadPage();
    }

    public function export1()
    {
        $result1 = Session::get('excel1');
        $fname = Session::get('fname');
        $title = Session::get('title');
        Excel::create($fname, function ($excel) use ($result1, $title) {
            $excel->sheet($title, function ($sheet) use ($result1) {
                $sheet->rows($result1);
                $sheet->freezeFirstRow();
            });
        })->export('xlsx');
    }

    public function export2()
    {
        $result2 = Session::get('excel2');
        $fname = Session::get('fname');
        $title = Session::get('title');
        $excel2 = Excel::create($fname . '-pdf', function ($excel) use ($result2, $title) {
            $excel->sheet($title, function ($sheet) use ($result2) {
                $rno = 0;
                $gno = 0;
                $bno = 0;
                $j = 0;
                $BStyle = array(
                    'borders' => array(
                        'outline' => array(
                            'style' => 'thin'
                        )
                    )
                );
                for ($i = 0; $i < 80; $i++) {
                    if ($i % 16 == 0 && $i != 0) {
                        $j++;
                    }
                    $sheet->mergeCells($sheet->getCellByColumnAndRow($i * 3 + 1 + 1 + $j, 15)->getCoordinate() . ":" . $sheet->getCellByColumnAndRow($i * 3 + 3 + 1 + $j, 15)->getCoordinate());
                    $sheet->mergeCells($sheet->getCellByColumnAndRow($i * 3 + 1 + 1 + $j, 1 + 15)->getCoordinate() . ":" . $sheet->getCellByColumnAndRow($i * 3 + 3 + 1 + $j, 1 + 15)->getCoordinate());
                    $sheet->mergeCells($sheet->getCellByColumnAndRow($i * 3 + 1 + 1 + $j, 29 + 15)->getCoordinate() . ":" . $sheet->getCellByColumnAndRow($i * 3 + 3 + 1 + $j, 29 + 15)->getCoordinate());
                    $sheet->mergeCells($sheet->getCellByColumnAndRow($i * 3 + 1 + 1 + $j, 57 + 15)->getCoordinate() . ":" . $sheet->getCellByColumnAndRow($i * 3 + 3 + 1 + $j, 57 + 15)->getCoordinate());
                    for ($z = 0; $z < 27; $z++) {
                        $sheet->getStyle($sheet->getCellByColumnAndRow($i * 3 + 1 + 1 + $j, 1 + 15 + 1 + $z)->getCoordinate() . ":" . $sheet->getCellByColumnAndRow($i * 3 + 3 + 1 + $j, 1 + 15 + 1 + $z)->getCoordinate())->applyFromArray($BStyle);
                        $sheet->getStyle($sheet->getCellByColumnAndRow($i * 3 + 1 + 1 + $j, 29 + 15 + 1 + $z)->getCoordinate() . ":" . $sheet->getCellByColumnAndRow($i * 3 + 3 + 1 + $j, 29 + 15 + 1 + $z)->getCoordinate())->applyFromArray($BStyle);
                    }
                    for ($y = 0; $y < 26; $y++) {
                        $sheet->getStyle($sheet->getCellByColumnAndRow($i * 3 + 1 + 1 + $j, 57 + 15 + 1 + $y)->getCoordinate() . ":" . $sheet->getCellByColumnAndRow($i * 3 + 3 + 1 + $j, 57 + 15 + 1 + $y)->getCoordinate())->applyFromArray($BStyle);
                    }
                }
                $sheet->row(15, array("", "", 1, "", "", 2, "", "", 3, "", "", 4, "", "", 5, "", "", 6, "", "", 7, "", "", 8, "", "", 9, "", "", 10, "", ""
                , 11, "", "", 12, "", "", 13, "", "", 14, "", "", 15, "", "", 16, "", "", "", 17, "", "", 18, "", "", 19, "", "", 20, "", "", 21, "", "", 22
                , "", "", 23, "", "", 24, "", "", 25, "", "", 26, "", "", 27, "", "", 28, "", "", 29, "", "", 30, "", "", 31, "", "", 32, "", "", "", 33, "", "",
                    34, "", "", 35, "", "", 36, "", "", 37, "", "", 38, "", "", 39, "", "", 40, "", "", 41, "", "", 42, "", "", 43, "", "", 44, "", "", 45, "", "",
                    46, "", "", 47, "", "", 48, "", "", "", 49, "", "", 50, "", "", 51, "", "", 52, "", "", 53, "", "", 54, "", "", 55, "", "", 56, "", "", 57, "", "",
                    58, "", "", 59, "", "", 60, "", "", 61, "", "", 62, "", "", 63, "", "", 64, "", "", "", 65, "", "", 66, "", "", 67, "", "", 68, "", "", 69, "", "",
                    70, "", "", 71, "", "", 72, "", "", 73, "", "", 74, "", "", 75, "", "", 76, "", "", 77, "", "", 78, "", "", 79, "", "", 80));
                $sheet->row(1 + 15, array("", "", 1, "", "", 2, "", "", 3, "", "", 4, "", "", 5, "", "", 6, "", "", 7, "", "", 8, "", "", 9, "", "", 10, "", "", 11, "", "", 12,
                    "", "", 13, "", "", 14, "", "", 15, "", "", 16, "", "", "", 1, "", "", 2, "", "", 3, "", "", 4, "", "", 5, "", "", 6, "", "", 7, "", "", 8, "",
                    "", 9, "", "", 10, "", "", 11, "", "", 12, "", "", 13, "", "", 14, "", "", 15, "", "", 16, "", "", "", 1, "", "", 2, "", "", 3, "", "", 4, "", "",
                    5, "", "", 6, "", "", 7, "", "", 8, "", "", 9, "", "", 10, "", "", 11, "", "", 12, "", "", 13, "", "",
                    14, "", "", 15, "", "", 16, "", "", "", 1, "", "", 2, "", "", 3, "", "", 4, "", "", 5, "", "", 6, "", "", 7, "", "",
                    8, "", "", 9, "", "", 10, "", "", 11, "", "", 12, "", "", 13, "", "", 14, "", "", 15, "", "", 16, "", "", "", 1, "", "", 2, "", "", 3, "", "",
                    4, "", "", 5, "", "", 6, "", "", 7, "", "", 8, "", "", 9, "", "", 10, "", "", 11, "", "", 12, "", "", 13, "", "", 14, "", "", 15, "", "", 16));
                $sheet->row(29 + 15, array("", "", 1, "", "", 2, "", "", 3, "", "", 4, "", "", 5, "", "", 6, "", "", 7, "", "", 8, "", "", 9, "", "", 10, "", "", 11, "", "", 12,
                    "", "", 13, "", "", 14, "", "", 15, "", "", 16, "", "", "", 1, "", "", 2, "", "", 3, "", "", 4, "", "", 5, "", "", 6, "", "", 7, "", "", 8, "",
                    "", 9, "", "", 10, "", "", 11, "", "", 12, "", "", 13, "", "", 14, "", "", 15, "", "", 16, "", "", "", 1, "", "", 2, "", "", 3, "", "", 4, "", "",
                    5, "", "", 6, "", "", 7, "", "", 8, "", "", 9, "", "", 10, "", "", 11, "", "", 12, "", "", 13, "", "",
                    14, "", "", 15, "", "", 16, "", "", "", 1, "", "", 2, "", "", 3, "", "", 4, "", "", 5, "", "", 6, "", "", 7, "", "",
                    8, "", "", 9, "", "", 10, "", "", 11, "", "", 12, "", "", 13, "", "", 14, "", "", 15, "", "", 16, "", "", "", 1, "", "", 2, "", "", 3, "", "",
                    4, "", "", 5, "", "", 6, "", "", 7, "", "", 8, "", "", 9, "", "", 10, "", "", 11, "", "", 12, "", "", 13, "", "", 14, "", "", 15, "", "", 16));
                $sheet->row(57 + 15, array("", "", 1, "", "", 2, "", "", 3, "", "", 4, "", "", 5, "", "", 6, "", "", 7, "", "", 8, "", "", 9, "", "", 10, "", "", 11, "", "", 12,
                    "", "", 13, "", "", 14, "", "", 15, "", "", 16, "", "", "", 1, "", "", 2, "", "", 3, "", "", 4, "", "", 5, "", "", 6, "", "", 7, "", "", 8, "",
                    "", 9, "", "", 10, "", "", 11, "", "", 12, "", "", 13, "", "", 14, "", "", 15, "", "", 16, "", "", "", 1, "", "", 2, "", "", 3, "", "", 4, "", "",
                    5, "", "", 6, "", "", 7, "", "", 8, "", "", 9, "", "", 10, "", "", 11, "", "", 12, "", "", 13, "", "",
                    14, "", "", 15, "", "", 16, "", "", "", 1, "", "", 2, "", "", 3, "", "", 4, "", "", 5, "", "", 6, "", "", 7, "", "",
                    8, "", "", 9, "", "", 10, "", "", 11, "", "", 12, "", "", 13, "", "", 14, "", "", 15, "", "", 16, "", "", "", 1, "", "", 2, "", "", 3, "", "",
                    4, "", "", 5, "", "", 6, "", "", 7, "", "", 8, "", "", 9, "", "", 10, "", "", 11, "", "", 12, "", "", 13, "", "", 14, "", "", 15, "", "", 16));


                for ($i = 1; $i <= 14; $i++) {
                    if ($i == 6 || $i == 7 || $i == 8) {
                        $sheet->setHeight($i, 85);
                    } else {
                        $sheet->setHeight($i, 45);
                    }
                }

                for ($i = 3; $i <= 246; $i++) {
                    if ($this->getNameFromNumber($i) == 'B' || $this->getNameFromNumber($i) == 'AY' || $this->getNameFromNumber($i) == 'CV'
                        || $this->getNameFromNumber($i) == 'ES' || $this->getNameFromNumber($i) == 'GP') {
                        $sheet->setWidth($this->getNameFromNumber($i), 7);
                    } else {
                        $sheet->setWidth($this->getNameFromNumber($i), 3);
                    }
                }
                for ($i = 0; $i < 27; $i++) {
                    $sheet->getCell('A' . ($i + 2 + 15))->setValue($i + 1);
                    $sheet->getCell('B' . ($i + 2 + 15))->setValue($i + 1);
                    $sheet->getCell('AY' . ($i + 2 + 15))->setValue($i + 1);
                    $sheet->getCell('CV' . ($i + 2 + 15))->setValue($i + 1);
                    $sheet->getCell('ES' . ($i + 2 + 15))->setValue($i + 1);
                    $sheet->getCell('GP' . ($i + 2 + 15))->setValue($i + 1);
                    $sheet->setHeight($i + 2 + 15, 45);
                }
                for ($i = 0; $i < 27; $i++) {
                    $sheet->getCell('A' . ($i + 2 + 28 + 15))->setValue($i + 1 + 27);
                    $sheet->getCell('B' . ($i + 2 + 28 + 15))->setValue($i + 1);
                    $sheet->getCell('AY' . ($i + 2 + 28 + 15))->setValue($i + 1);
                    $sheet->getCell('CV' . ($i + 2 + 28 + 15))->setValue($i + 1);
                    $sheet->getCell('ES' . ($i + 2 + 28 + 15))->setValue($i + 1);
                    $sheet->getCell('GP' . ($i + 2 + 28 + 15))->setValue($i + 1);
                    $sheet->setHeight($i + 2 + 28 + 15, 45);
                }
                for ($i = 0; $i < 26; $i++) {
                    $sheet->getCell('A' . ($i + 2 + 56 + 15))->setValue($i + 1 + 54);
                    $sheet->getCell('B' . ($i + 2 + 56 + 15))->setValue($i + 1);
                    $sheet->getCell('AY' . ($i + 2 + 56 + 15))->setValue($i + 1);
                    $sheet->getCell('CV' . ($i + 2 + 56 + 15))->setValue($i + 1);
                    $sheet->getCell('ES' . ($i + 2 + 56 + 15))->setValue($i + 1);
                    $sheet->getCell('GP' . ($i + 2 + 56 + 15))->setValue($i + 1);
                    $sheet->setHeight($i + 2 + 56 + 15, 45);
                }


                if (!empty($result2)) {
                    foreach ($result2 as $v) {
                        if ($v['group'] <= 10) {
                            if ($v['group'] % 5 == 0) {//group 5.10
                                if ($v['color'] == "R") {
                                    $rno++;
                                    $sheet->cell($sheet->getCellByColumnAndRow(4 * 16 * 3 + ($v['lines'] - 1) * 3 + 1 + 4 + 1, (ceil($v['group'] / 5) - 1) * 27 + $v['rows'] + ceil($v['group'] / 5) + 15)->getCoordinate(), function ($cell) use ($v) {
                                        $cell->setBackground("#FF0000");
                                    });
                                } elseif ($v['color'] == "G") {
                                    $gno++;
                                    $sheet->cell($sheet->getCellByColumnAndRow(4 * 16 * 3 + ($v['lines'] - 1) * 3 + 2 + 4 + 1, (ceil($v['group'] / 5) - 1) * 27 + $v['rows'] + ceil($v['group'] / 5) + 15)->getCoordinate(), function ($cell) use ($v) {
                                        $cell->setBackground("#00FF11");
                                    });
                                } elseif ($v['color'] == "B") {
                                    $bno++;
                                    $sheet->cell($sheet->getCellByColumnAndRow(4 * 16 * 3 + ($v['lines'] - 1) * 3 + 3 + 4 + 1, (ceil($v['group'] / 5) - 1) * 27 + $v['rows'] + ceil($v['group'] / 5) + 15)->getCoordinate(), function ($cell) use ($v) {
                                        $cell->setBackground("#002BFF");
                                    });
                                }
                            } else { //group 1.2.3.4.6.7.8.9
                                if ($v['color'] == "R") {
                                    $rno++;
                                    $sheet->cell($sheet->getCellByColumnAndRow(
                                        (($v['group'] % 5) - 1) * 16 * 3 + ($v['lines'] - 1) * 3 + 1 + ($v['group'] % 5) - 1 + 1,
                                        (ceil($v['group'] / 5) - 1) * 27 + $v['rows'] + ceil($v['group'] / 5) + 15
                                    )
                                        ->getCoordinate(), function ($cell) use ($v) {
                                        $cell->setBackground("#FF0000");
                                    });
                                } elseif ($v['color'] == "G") {
                                    $gno++;
                                    $sheet->cell($sheet->getCellByColumnAndRow(
                                        (($v['group'] % 5) - 1) * 16 * 3 + ($v['lines'] - 1) * 3 + 2 + ($v['group'] % 5) - 1 + 1,
                                        (ceil($v['group'] / 5) - 1) * 27 + $v['rows'] + ceil($v['group'] / 5) + 15
                                    )
                                        ->getCoordinate(), function ($cell) use ($v) {
                                        $cell->setBackground("#00FF11");
                                    });
                                } elseif ($v['color'] == "B") {
                                    $bno++;
                                    $sheet->cell($sheet->getCellByColumnAndRow(
                                        (($v['group'] % 5) - 1) * 16 * 3 + ($v['lines'] - 1) * 3 + 3 + ($v['group'] % 5) - 1 + 1,
                                        (ceil($v['group'] / 5) - 1) * 27 + $v['rows'] + ceil($v['group'] / 5) + 15
                                    )
                                        ->getCoordinate(), function ($cell) use ($v) {
                                        $cell->setBackground("#002BFF");
                                    });
                                }
                            }
                        } else {
                            if ($v['group'] % 5 == 0) {//group 15
                                if ($v['color'] == "R") {
                                    $rno++;
                                    $sheet->cell($sheet->getCellByColumnAndRow(4 * 16 * 3 + ($v['lines'] - 1) * 3 + 1 + 4 + 1, 54 + $v['rows'] + ceil($v['group'] / 5) + 15)->getCoordinate(), function ($cell) use ($v) {
                                        $cell->setBackground("#FF0000");
                                    });
                                } elseif ($v['color'] == "G") {
                                    $gno++;
                                    $sheet->cell($sheet->getCellByColumnAndRow(4 * 16 * 3 + ($v['lines'] - 1) * 3 + 2 + 4 + 1, 54 + $v['rows'] + ceil($v['group'] / 5) + 15)->getCoordinate(), function ($cell) use ($v) {
                                        $cell->setBackground("#00FF11");
                                    });
                                } elseif ($v['color'] == "B") {
                                    $bno++;
                                    $sheet->cell($sheet->getCellByColumnAndRow(4 * 16 * 3 + ($v['lines'] - 1) * 3 + 3 + 4 + 1, 54 + $v['rows'] + ceil($v['group'] / 5) + 15)->getCoordinate(), function ($cell) use ($v) {
                                        $cell->setBackground("#002BFF");
                                    });
                                }
                            } else { //group 11.12.13.14

                                if ($v['color'] == "R") {
                                    $rno++;
                                    $sheet->cell($sheet->getCellByColumnAndRow((($v['group'] % 5) - 1) * 16 * 3 + ($v['lines'] - 1) * 3 + 1 + ($v['group'] % 5) - 1 + 1, 54 + $v['rows'] + ceil($v['group'] / 5) + 15)
                                        ->getCoordinate(), function ($cell) use ($v) {
                                        $cell->setBackground("#FF0000");
                                    });
                                } elseif ($v['color'] == "G") {
                                    $gno++;
                                    $sheet->cell($sheet->getCellByColumnAndRow((($v['group'] % 5) - 1) * 16 * 3 + ($v['lines'] - 1) * 3 + 2 + ($v['group'] % 5) - 1 + 1, 54 + $v['rows'] + ceil($v['group'] / 5) + 15)
                                        ->getCoordinate(), function ($cell) use ($v) {
                                        $cell->setBackground("#00FF11");
                                    });
                                } elseif ($v['color'] == "B") {
                                    $bno++;
                                    $sheet->cell($sheet->getCellByColumnAndRow((($v['group'] % 5) - 1) * 16 * 3 + ($v['lines'] - 1) * 3 + 3 + ($v['group'] % 5) - 1 + 1, 54 + $v['rows'] + ceil($v['group'] / 5) + 15)
                                        ->getCoordinate(), function ($cell) use ($v) {
                                        $cell->setBackground("#002BFF");
                                    });
                                }
                            }
                        }
                    }
                } else {
                    $sheet->rows($result2);
                }
                //printdate
                $sheet->mergeCells('D3:W5');
                $sheet->getCell('D3')->setValue('Printed Date:');//printdate cell
                $sheet->cells('D3:W5', function ($cells) {
                    $cells->setFont(
                        array(
                            'family' => 'Calibri',
                            'size' => '68',
                        )
                    );
                    $cells->setAlignment('center');
                    $cells->setValignment('center');
                });
                $sheet->setBorder('D3', 'thin');
                //printdate input
                $sheet->mergeCells('X3:BO5');
                $sheet->setBorder('X3', 'thin');
                //R
                $sheet->mergeCells('D8:G8');
                $sheet->getCell('D8')->setValue('R');
                $sheet->cells('D8:G8', function ($cells) {
                    $cells->setFont(
                        array(
                            'family' => 'Calibri',
                            'size' => '72',
                        )
                    );
                    $cells->setAlignment('center');
                    $cells->setValignment('center');
                });
                $sheet->setBorder('D8', 'thin');
                //R color
                $sheet->mergeCells('D9:G11');
                $sheet->cells('D9:G11', function ($cells) {
                    $cells->setBackground("#888888");
                });
                $sheet->setBorder('D9', 'thin');
                //G
                $sheet->mergeCells('H8:K8');
                $sheet->getCell('H8')->setValue('G');
                $sheet->cells('H8:K8', function ($cells) {
                    $cells->setFont(
                        array(
                            'family' => 'Calibri',
                            'size' => '72',
                        )
                    );
                    $cells->setAlignment('center');
                    $cells->setValignment('center');
                });
                $sheet->setBorder('H8', 'thin');
                //G Color
                $sheet->mergeCells('H9:K11');
                $sheet->cells('H9:K11', function ($cells) {
                    $cells->setBackground("#666666");
                });
                $sheet->setBorder('H9', 'thin');
                //B
                $sheet->mergeCells('L8:O8');
                $sheet->getCell('L8')->setValue('B');
                $sheet->cells('L8:O8', function ($cells) {
                    $cells->setFont(
                        array(
                            'family' => 'Calibri',
                            'size' => '72',
                        )
                    );
                    $cells->setAlignment('center');
                    $cells->setValignment('center');
                });
                $sheet->setBorder('L8', 'thin');
                //B Color
                $sheet->mergeCells('L9:O11');
                $sheet->cells('L9:O11', function ($cells) {
                    $cells->setBackground("#444444");
                });
                $sheet->setBorder('L9', 'thin');
                //lot no
                $sheet->mergeCells('CE3:CR5');
                $sheet->getCell('CE3')->setValue('Lot No:');
                $sheet->setBorder('CE3', 'thin');
                $sheet->cells('CE3:CR5', function ($cells) {
                    $cells->setFont(
                        array(
                            'family' => 'Calibri',
                            'size' => '72',
                        )
                    );
                    $cells->setAlignment('center');
                    $cells->setValignment('center');
                });
                //lot no input
                $sheet->mergeCells('CS3:EM5');
                $sheet->setBorder('CS3', 'thin');
                //L板
                $sheet->mergeCells('CE6:EM6');
                $sheet->getCell('CE6')->setValue('L板Bonding 需180度倒轉');
                $sheet->setBorder('CE6', 'thin');
                $sheet->cells('CE6:EM6', function ($cells) {
                    $cells->setFont(
                        array(
                            'family' => 'Calibri',
                            'size' => '72',
                        )
                    );
                    $cells->setAlignment('center');
                    $cells->setValignment('center');
                });
                //name
                $sheet->mergeCells('FA3:FO5');
                $sheet->getCell('FA3')->setValue('Name:');
                $sheet->setBorder('FA3', 'thin');
                $sheet->cells('FA3:FO5', function ($cells) {
                    $cells->setFont(
                        array(
                            'family' => 'Calibri',
                            'size' => '72',
                        )
                    );
                    $cells->setAlignment('center');
                    $cells->setValignment('center');
                });
                //name input
                $sheet->mergeCells('FP3:GK5');
                $sheet->setBorder('FP3', 'thin');
                //date
                $sheet->mergeCells('FA6:FO7');
                $sheet->getCell('FA6')->setValue('Date:');
                $sheet->setBorder('FA6', 'thin');
                $sheet->cells('FA6:FO7', function ($cells) {
                    $cells->setFont(
                        array(
                            'family' => 'Calibri',
                            'size' => '72',
                        )
                    );
                    $cells->setAlignment('center');
                    $cells->setValignment('center');
                });
                //date input
                $sheet->mergeCells('FP6:GK7');
                $sheet->setBorder('FP6', 'thin');
                //defect record
                $sheet->mergeCells('HG3:IA5');
                $sheet->getCell('HG3')->setValue('Defect record');
                $sheet->setBorder('HG3', 'thin');
                $sheet->cells('HG3:IA5', function ($cells) {
                    $cells->setFont(
                        array(
                            'family' => 'Calibri',
                            'size' => '68',
                        )
                    );
                    $cells->setAlignment('center');
                    $cells->setValignment('center');
                });
                //R
                $sheet->mergeCells('HG6:HK6');
                $sheet->getCell('HG6')->setValue('R:');
                $sheet->setBorder('HG6', 'thin');
                //R input
                $sheet->mergeCells('HL6:IA6');
                $sheet->getCell('HL6')->setValue($rno);
                $sheet->setBorder('HL6', 'thin');
                //G
                $sheet->mergeCells('HG7:HK7');
                $sheet->getCell('HG7')->setValue('G:');
                $sheet->setBorder('HG7', 'thin');
                //G input
                $sheet->mergeCells('HL7:IA7');
                $sheet->getCell('HL7')->setValue($gno);
                $sheet->setBorder('HL7', 'thin');
                //B
                $sheet->mergeCells('HG8:HK8');
                $sheet->getCell('HG8')->setValue('B:');
                $sheet->setBorder('HG8', 'thin');
                //B input
                $sheet->mergeCells('HL8:IA8');
                $sheet->getCell('HL8')->setValue($bno);
                $sheet->setBorder('HL8', 'thin');

                $sheet->cells('HG6:IA8', function ($cells) {
                    $cells->setFont(
                        array(
                            'family' => 'Calibri',
                            'size' => '63',
                        )
                    );
                    $cells->setAlignment('center');
                    $cells->setValignment('center');
                });

                $sheet->cells('A15:IL15', function ($cells) {
                    $cells->setFont(
                        array(
                            'family' => 'Calibri',
                            'size' => '28',
                            'bold' => true
                        )
                    );
                    $cells->setAlignment('center');
                    $cells->setValignment('center');
                });
                $sheet->cells('A17:A98', function ($cells) {
                    $cells->setFont(
                        array(
                            'family' => 'Calibri',
                            'size' => '28',
                            'bold' => true
                        )
                    );
                    $cells->setAlignment('center');
                    $cells->setValignment('center');
                });


                $sheet->cells('B16:IL98', function ($cells) {
                    $cells->setFont(
                        array(
                            'family' => 'Calibri',
                            'size' => '28',
                        )
                    );
                    $cells->setAlignment('center');
                    $cells->setValignment('center');
                });

                $sheet->getsheetView()->setzoomScale('25');
                $sheet->getsheetView()->setzoomScaleNormal('25');
            });
        })->export('xlsx');
    }

    public function export3()
    {
        $result2 = Session::get('excel2');
        $chip = Session::get('chip');
        $pixel = Session::get('pixel');
        $fname = Session::get('fname');
        $content = "";
        $result3 = array();
        foreach ($result2 as $v) {
            if ($v['color'] == "R") {
                $l = $pixel + $chip;
            } elseif ($v['color'] == "G") {
                $l = $pixel;
            } elseif ($v['color'] == "B") {
                $l = $pixel - $chip;
            }
            if ($v['group'] <= 10) {//1.2.3.4.5.6.7.8.9.10
                if ($v['group'] % 5 == 0) {//5.10
                    $arr = ["color" => $v['color'], "x" => (-1) * $pixel * 16 * 4 + $v['lines'] * (-1) * ($pixel) + $pixel,
                        "y" => (-1) * $pixel * 27 * (ceil($v['group'] / 5) - 1) + $v['rows'] * (-1) * ($pixel) + $l];
                    array_push($result3, $arr);
                } else { //1.2.3.4.6.7.8.9
                    $arr = ["color" => $v['color'], "x" => (-1) * ($pixel) * 16 * ($v['group'] % 5 - 1) + (-1) * ($pixel) * $v['lines'] + $pixel,
                        "y" => (-1) * $pixel * 27 * (ceil($v['group'] / 5) - 1) + $v['rows'] * (-1) * ($pixel) + $l];
                    array_push($result3, $arr);
                }
            } else { //11.12.13.14.15
                if ($v['group'] % 5 == 0) {//15
                    $arr = ["color" => $v['color'], "x" => (-1) * ($pixel) * 16 * 4 + $v['lines'] * (-1) * ($pixel) + $pixel,
                        "y" => (-1) * $pixel * 54 + $v['rows'] * (-1) * ($pixel) + $l];
                    array_push($result3, $arr);
                } else { //11.12.13.14
                    $arr = ["color" => $v['color'], "x" => (-1) * ($pixel) * 16 * ($v['group'] % 5 - 1) + (-1) * ($pixel) * $v['lines'] + $pixel,
                        "y" => (-1) * $pixel * 54 + $v['rows'] * (-1) * ($pixel) + $l];
                    array_push($result3, $arr);
                }
            }
        }
        usort($result3, $this->arrSortObjsByKey('y', 'ASC'));
        foreach ($result3 as $v) {
            $content .= $v['color'] . "\t" . $v['x'] . "\t" . $v['y'];
            $content .= "\r\n";
        }
        Storage::put($fname, $content);
        return response()->download(storage_path('app/' . $fname), $fname . '.Lxy')->deleteFileAfterSend(true);;
    }



    public function export4R()
    {
        $result1 = Session::get('excel1');
        $result4 = array();
        $content = '';
        $fname = Session::get('fname') . "-R";

        unset($result1[0]);
        foreach ($result1 as $v) {
            $array = array(
                "rows" => $v['group']<10?(ceil($v['group'] / 5) - 1) * (27) + $v['rows']:54+$v['rows'],
                "lines" => ($v['group'] - 1) % 5 * (16) + $v['lines'],
                "color" => $v['color'],
                "result" => $v['result']
            );
            array_push($result4, $array);
        }

        for ($i = 1; $i <= 80; $i--) {
            for ($j = 1; $j <= 80; $j++) {
                foreach ($result4 as $v) {
                    if ($v['rows'] == $i && $v['lines'] == $j) {
                        if ($v['color'] == 'R') {
                            $content .= ($v['result'] == 'PASS' ? 1 : 0) . ",";
                        }
                    }
                }
            }
        }
        Storage::put($fname, $content);
        return response()->download(storage_path('app/' . $fname), $fname . '.mpd')->deleteFileAfterSend(true);

    }

    public function export4G()
    {
        $result1 = Session::get('excel1');
        $result4 = array();
        $content = '';
        $fname = Session::get('fname') . "-G";

        unset($result1[0]);
        foreach ($result1 as $v) {
            $array = array(
                "rows" => $v['group']<10?(ceil($v['group'] / 5) - 1) * (27) + $v['rows']:54+$v['rows'],
                "lines" => ($v['group'] - 1) % 5 * (16) + $v['lines'],
                "color" => $v['color'],
                "result" => $v['result']
            );
            array_push($result4, $array);
        }


        for ($i = 1; $i <= 80; $i--) {
            for ($j = 1; $j <= 80; $j++) {
                foreach ($result4 as $v) {
                    if ($v['rows'] == $i && $v['lines'] == $j) {
                        if ($v['color'] == 'R') {
                            $content .= ($v['result'] == 'PASS' ? 1 : 0) . ",";
                        }
                    }
                }
            }
        }
        Storage::put($fname, $content);
        return response()->download(storage_path('app/' . $fname), $fname . '.mpd')->deleteFileAfterSend(true);

    }

    public function export4B()
    {
        $result1 = Session::get('excel1');
        $result4 = array();
        $content = '';
        $fname = Session::get('fname') . "-B";

        unset($result1[0]);
        foreach ($result1 as $v) {
            $array = array(
                "rows" => $v['group']<10?(ceil($v['group'] / 5) - 1) * (27) + $v['rows']:54+$v['rows'],
                "lines" => ($v['group'] - 1) % 5 * (16) + $v['lines'],
                "color" => $v['color'],
                "result" => $v['result']
            );
            array_push($result4, $array);
        }


        for ($i = 1; $i <= 80; $i--) {
            for ($j = 1; $j <= 80; $j++) {
                foreach ($result4 as $v) {
                    if ($v['rows'] == $i && $v['lines'] == $j) {
                        if ($v['color'] == 'R') {
                            $content .= ($v['result'] == 'PASS' ? 1 : 0) . ",";
                        }
                    }
                }
            }
        }

        Storage::put($fname, $content);
        return response()->download(storage_path('app/' . $fname), $fname . '.mpd')->deleteFileAfterSend(true);

    }


    public function median($numbers = array())
    {
        if (!is_array($numbers)) {
            $numbers = func_get_args();
        }
        rsort($numbers);
        $mid = (count($numbers) / 2);
        return ($mid % 2 != 0) ? $numbers{$mid - 1} : (($numbers{$mid - 1}) + $numbers{$mid}) / 2;
    }

    public function getNameFromNumber($num)
    {
        $numeric = ($num - 1) % 26;
        $letter = chr(65 + $numeric);
        $num2 = intval(($num - 1) / 26);
        if ($num2 > 0) {
            return $this->getNameFromNumber($num2) . $letter;
        } else {
            return $letter;
        }
    }

    function arrSortObjsByKey($key, $order = 'DESC')
    {
        return function ($a, $b) use ($key, $order) {
            // Swap order if necessary
            if ($order == 'DESC') {
                list($a, $b) = array($b, $a);
            }
            // Check data type
            if (is_numeric($a[$key])) {
                return $a[$key] - $b[$key]; // compare numeric
            } else {
                return strnatcasecmp($a[$key], $b[$key]); // compare string
            }
        };
    }
}
