<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Excel;


class ExcelController extends Controller
{

  public function getViewPage()
  {
    return view('excel');
  }
  public function upload(Request $request){
    $filepath=$request->file('file')->getRealPath();
    $rupper=$request->input('rupper');
    $rlower=$request->input('rlower')*(-1);
    $rstd=$request->input('rstandard');
    $gupper=$request->input('gupper');
    $glower=$request->input('glower')*(-1);
    $gstd=$request->input('gstandard');
    $bupper=$request->input('bupper');
    $blower=$request->input('blower')*(-1);
    $bstd=$request->input('bstandard');
    $excel=Excel::load($filepath)->get();
    $odata=$excel->toArray();
    $title=$excel->getTitle();
    $result1=array();
    $result2=array();
    $rmvarr=array();//red measure value
    $gmvarr=array();//green measure value
    $bmvarr=array();//blue measure value
    $chip=$request->input('chip');
    $pixel=$request->input('pixel');
    foreach ($odata as $v) {
        if($v['color']=="R"){
          array_push($rmvarr,str_replace(" V","",$v['measure_value']));
        }
        else if($v['color']=="G"){
          array_push($gmvarr,str_replace(" V","",$v['measure_value']));
        }
        else if($v['color']=="B"){
          array_push($bmvarr,str_replace(" V","",$v['measure_value']));
        }
    }
    if(!$rstd)
    {
      $rstd=$this->median($rmvarr);
    }
    if(!$gstd)
    {
      $gstd=$this->median($gmvarr);
    }
    if(!$bstd)
    {
      $bstd=$this->median($bmvarr);
    }
    $rmin=$rstd*(1+$rlower/100);
    $rmax=$rstd*(1+$rupper/100);
    $gmin=$gstd*(1+$glower/100);
    $gmax=$gstd*(1+$gupper/100);
    $bmin=$bstd*(1+$blower/100);
    $bmax=$bstd*(1+$bupper/100);

    array_push($result1,['Group','Rows','Lines','Color','Standard Value','Hi Limit','Low Limit','Measure Value','Result']);
    foreach ($odata as $v) {
      if($v['color']=="R"){
        if(str_replace(" V","",$v['measure_value'])<=$rmax && str_replace(" V","",$v['measure_value'])>=$rmin)//pass
        {
          $array = array(
              "group"    => $v['group'],
              "rows"  => $v['rows'],
              "lines" => $v['lines'],
              "color" => $v['color'],
              "standard_value"=>$rstd,
              "hi_limit"=>$rmax,
              "low_limit"=>$rmin,
              "measure_value"=>str_replace(" V","",$v['measure_value']),
              "result"=>"PASS"
          );
        }
        else {//ng
          $array = array(
              "group"    => $v['group'],
              "rows"  => $v['rows'],
              "lines" => $v['lines'],
              "color" => $v['color'],
              "standard_value"=>$rstd,
              "hi_limit"=>$rmax,
              "low_limit"=>$rmin,
              "measure_value"=>str_replace(" V","",$v['measure_value']),
              "result"=>"NG"
          );
          $ngplace=array(
            "group"    => $v['group'],
            "rows"  => $v['rows'],
            "lines" => $v['lines'],
            "color" => $v['color']
          );
          array_push($result2,$ngplace);
        }
        array_push($result1,$array);
      }
      else if($v['color']=="G"){
        if(str_replace(" V","",$v['measure_value'])<=$gmax && str_replace(" V","",$v['measure_value'])>=$gmin)//pass
        {
          $array = array(
              "group"    => $v['group'],
              "rows"  => $v['rows'],
              "lines" => $v['lines'],
              "color" => $v['color'],
              "standard_value"=>$gstd,
              "hi_limit"=>$gmax,
              "low_limit"=>$gmin,
              "measure_value"=>str_replace(" V","",$v['measure_value']),
              "result"=>"PASS"
          );
        }
        else {//ng
          $array = array(
              "group"    => $v['group'],
              "rows"  => $v['rows'],
              "lines" => $v['lines'],
              "color" => $v['color'],
              "standard_value"=>$gstd,
              "hi_limit"=>$gmax,
              "low_limit"=>$gmin,
              "measure_value"=>str_replace(" V","",$v['measure_value']),
              "result"=>"NG"
          );
          $ngplace=array(
            "group"    => $v['group'],
            "rows"  => $v['rows'],
            "lines" => $v['lines'],
            "color" => $v['color']
          );
          array_push($result2,$ngplace);
        }
        array_push($result1,$array);
      }
      else if($v['color']=="B"){
        if(str_replace(" V","",$v['measure_value'])<=$bmax && str_replace(" V","",$v['measure_value'])>=$bmin)//pass
        {
          $array = array(
              "group"    => $v['group'],
              "rows"  => $v['rows'],
              "lines" => $v['lines'],
              "color" => $v['color'],
              "standard_value"=>$bstd,
              "hi_limit"=>$bmax,
              "low_limit"=>$bmin,
              "measure_value"=>str_replace(" V","",$v['measure_value']),
              "result"=>"PASS"
          );
        }
        else {//ng
          $array = array(
              "group"    => $v['group'],
              "rows"  => $v['rows'],
              "lines" => $v['lines'],
              "color" => $v['color'],
              "standard_value"=>$bstd,
              "hi_limit"=>$bmax,
              "low_limit"=>$bmin,
              "measure_value"=>str_replace(" V","",$v['measure_value']),
              "result"=>"NG"
          );
          $ngplace=array(
            "group"    => $v['group'],
            "rows"  => $v['rows'],
            "lines" => $v['lines'],
            "color" => $v['color']
          );
          array_push($result2,$ngplace);
        }
        array_push($result1,$array);
      }
    }
    $ex1=Excel::create('requirement1',function ($excel) use($title,$result1,$result2,$chip,$pixel){
        $excel->sheet($title."1", function ($sheet) use ($result1){
          $sheet->rows($result1);
          $sheet->freezeFirstRow();
        });
        $excel->sheet($title."2",function($sheet) use($result2){
          $rno=0;
          $gno=0;
          $bno=0;
          $j=0;
          $BStyle = array(
            'borders' => array(
              'outline' => array(
                'style' => 'thin'
              )
            )
          );
          for($i=0;$i<80;$i++)
          {
            if($i%16==0&&$i!=0)
            {
              $j++;
            }
            $sheet->mergeCells($sheet->getCellByColumnAndRow($i*3+1+$j,1+15)->getCoordinate().":".$sheet->getCellByColumnAndRow($i*3+3+$j,1+15)->getCoordinate());
            $sheet->mergeCells($sheet->getCellByColumnAndRow($i*3+1+$j,29+15)->getCoordinate().":".$sheet->getCellByColumnAndRow($i*3+3+$j,29+15)->getCoordinate());
            $sheet->mergeCells($sheet->getCellByColumnAndRow($i*3+1+$j,57+15)->getCoordinate().":".$sheet->getCellByColumnAndRow($i*3+3+$j,57+15)->getCoordinate());
            for($z=0;$z<27;$z++)
            {
              $sheet->getStyle($sheet->getCellByColumnAndRow($i*3+1+$j,1+15+1+$z)->getCoordinate().":".$sheet->getCellByColumnAndRow($i*3+3+$j,1+15+1+$z)->getCoordinate())->applyFromArray($BStyle);
              $sheet->getStyle($sheet->getCellByColumnAndRow($i*3+1+$j,29+15+1+$z)->getCoordinate().":".$sheet->getCellByColumnAndRow($i*3+3+$j,29+15+1+$z)->getCoordinate())->applyFromArray($BStyle);
            }
            for($y=0;$y<26;$y++)
            {
              $sheet->getStyle($sheet->getCellByColumnAndRow($i*3+1+$j,57+15+1+$y)->getCoordinate().":".$sheet->getCellByColumnAndRow($i*3+3+$j,57+15+1+$y)->getCoordinate())->applyFromArray($BStyle);
            }

          }
          $sheet->row(1+15,array("",1,"","",2,"","",3,"","",4,"","",5,"","",6,"","",7,"","",8,"","",9,"","",10,"","",11,"","",12,
          "","",13,"","",14,"","",15,"","",16,"","","",1,"","",2,"","",3,"","",4,"","",5,"","",6,"","",7,"","",8,"",
          "",9,"","",10,"","",11,"","",12,"","",13,"","",14,"","",15,"","",16,"","","",1,"","",2,"","",3,"","",4,"","",
          5,"","",6,"","",7,"","",8,"","",9,"","",10,"","",11,"","",12,"","",13,"","",
          14,"","",15,"","",16,"","","",1,"","",2,"","",3,"","",4,"","",5,"","",6,"","",7,"","",
          8,"","",9,"","",10,"","",11,"","",12,"","",13,"","",14,"","",15,"","",16,"","","",1,"","",2,"","",3,"","",
          4,"","",5,"","",6,"","",7,"","",8,"","",9,"","",10,"","",11,"","",12,"","",13,"","",14,"","",15,"","",16));
          $sheet->row(29+15,array("",1,"","",2,"","",3,"","",4,"","",5,"","",6,"","",7,"","",8,"","",9,"","",10,"","",11,"","",12,
          "","",13,"","",14,"","",15,"","",16,"","","",1,"","",2,"","",3,"","",4,"","",5,"","",6,"","",7,"","",8,"",
          "",9,"","",10,"","",11,"","",12,"","",13,"","",14,"","",15,"","",16,"","","",1,"","",2,"","",3,"","",4,"","",
          5,"","",6,"","",7,"","",8,"","",9,"","",10,"","",11,"","",12,"","",13,"","",
          14,"","",15,"","",16,"","","",1,"","",2,"","",3,"","",4,"","",5,"","",6,"","",7,"","",
          8,"","",9,"","",10,"","",11,"","",12,"","",13,"","",14,"","",15,"","",16,"","","",1,"","",2,"","",3,"","",
          4,"","",5,"","",6,"","",7,"","",8,"","",9,"","",10,"","",11,"","",12,"","",13,"","",14,"","",15,"","",16));
          $sheet->row(57+15,array("",1,"","",2,"","",3,"","",4,"","",5,"","",6,"","",7,"","",8,"","",9,"","",10,"","",11,"","",12,
          "","",13,"","",14,"","",15,"","",16,"","","",1,"","",2,"","",3,"","",4,"","",5,"","",6,"","",7,"","",8,"",
          "",9,"","",10,"","",11,"","",12,"","",13,"","",14,"","",15,"","",16,"","","",1,"","",2,"","",3,"","",4,"","",
          5,"","",6,"","",7,"","",8,"","",9,"","",10,"","",11,"","",12,"","",13,"","",
          14,"","",15,"","",16,"","","",1,"","",2,"","",3,"","",4,"","",5,"","",6,"","",7,"","",
          8,"","",9,"","",10,"","",11,"","",12,"","",13,"","",14,"","",15,"","",16,"","","",1,"","",2,"","",3,"","",
          4,"","",5,"","",6,"","",7,"","",8,"","",9,"","",10,"","",11,"","",12,"","",13,"","",14,"","",15,"","",16));




          for($i=0;$i<27;$i++)
          {
            $sheet->getCell('A'.($i+2+15))->setValue($i+1);
            $sheet->getCell('AX'.($i+2+15))->setValue($i+1);
            $sheet->getCell('CU'.($i+2+15))->setValue($i+1);
            $sheet->getCell('ER'.($i+2+15))->setValue($i+1);
            $sheet->getCell('GO'.($i+2+15))->setValue($i+1);
            $sheet->setHeight($i+2+15,45);
          }
          for($i=0;$i<27;$i++)
          {
            $sheet->getCell('A'.($i+2+28+15))->setValue($i+1);
            $sheet->getCell('AX'.($i+2+28+15))->setValue($i+1);
            $sheet->getCell('CU'.($i+2+28+15))->setValue($i+1);
            $sheet->getCell('ER'.($i+2+28+15))->setValue($i+1);
            $sheet->getCell('GO'.($i+2+28+15))->setValue($i+1);
            $sheet->setHeight($i+2+28+15,45);
          }
          for($i=0;$i<26;$i++)
          {
            $sheet->getCell('A'.($i+2+56+15))->setValue($i+1);
            $sheet->getCell('AX'.($i+2+56+15))->setValue($i+1);
            $sheet->getCell('CU'.($i+2+56+15))->setValue($i+1);
            $sheet->getCell('ER'.($i+2+56+15))->setValue($i+1);
            $sheet->getCell('GO'.($i+2+56+15))->setValue($i+1);
            $sheet->setHeight($i+2+56+15,45);
          }
          if(!empty($result2))
          {
            foreach($result2 as $v)
            {

              if($v['group']<=10)
              {
                  if($v['group']%5==0)//group 5.10
                  {
                    if($v['color']=="R")
                    {
                      $rno++;
                      $sheet->cell($sheet->getCellByColumnAndRow(4*16*3+($v['lines']-1)*3+1+4,(ceil($v['group']/5)-1)*27+$v['rows']+ceil($v['group']/5)+15)->getCoordinate(),function($cell) use($v){
                        $cell->setBackground("#FF0000");
                      });
                    }
                    else if($v['color']=="G")
                    {
                      $gno++;
                      $sheet->cell($sheet->getCellByColumnAndRow(4*16*3+($v['lines']-1)*3+2+4,(ceil($v['group']/5)-1)*27+$v['rows']+ceil($v['group']/5)+15)->getCoordinate(),function($cell) use($v){
                        $cell->setBackground("#00FF11");
                      });
                    }
                    else if($v['color']=="B")
                    {
                      $bno++;
                      $sheet->cell($sheet->getCellByColumnAndRow(4*16*3+($v['lines']-1)*3+3+4,(ceil($v['group']/5)-1)*27+$v['rows']+ceil($v['group']/5)+15)->getCoordinate(),function($cell) use($v){
                        $cell->setBackground("#002BFF");
                      });
                    }
                  }

                else { //group 1.2.3.4.6.7.8.9
                    if($v['color']=="R")
                    {
                      $rno++;
                      $sheet->cell($sheet->getCellByColumnAndRow((($v['group']%5)-1)*16*3+($v['lines']-1)*3+1+($v['group']%5)-1,
                      (ceil($v['group']/5)-1)*27+$v['rows']+ceil($v['group']/5)+15)
                      ->getCoordinate(),function($cell)
                      use($v){
                        $cell->setBackground("#FF0000");
                      });
                    }
                    else if($v['color']=="G")
                    {
                      $gno++;
                      $sheet->cell($sheet->getCellByColumnAndRow((($v['group']%5)-1)*16*3+($v['lines']-1)*3+2+($v['group']%5)-1,
                      (ceil($v['group']/5)-1)*27+$v['rows']+ceil($v['group']/5)+15)
                      ->getCoordinate(),function($cell)
                      use($v){
                        $cell->setBackground("#00FF11");
                      });
                    }
                    else if($v['color']=="B")
                    {
                      $bno++;
                      $sheet->cell($sheet->getCellByColumnAndRow((($v['group']%5)-1)*16*3+($v['lines']-1)*3+3+($v['group']%5)-1,
                      (ceil($v['group']/5)-1)*27+$v['rows']+ceil($v['group']/5)+15)
                      ->getCoordinate(),function($cell)
                      use($v){
                        $cell->setBackground("#002BFF");
                      });
                    }
                }
              }
              else {
                if($v['group']%5==0)//group 15
                {
                  if($v['color']=="R")
                  {
                    $rno++;
                    $sheet->cell($sheet->getCellByColumnAndRow(4*16*3+($v['lines']-1)*3+1+4,54+$v['rows']+ceil($v['group']/5)+15)->getCoordinate(),function($cell) use($v){
                      $cell->setBackground("#FF0000");
                    });
                  }
                  else if($v['color']=="G")
                  {
                    $gno++;
                    $sheet->cell($sheet->getCellByColumnAndRow(4*16*3+($v['lines']-1)*3+2+4,54+$v['rows']+ceil($v['group']/5)+15)->getCoordinate(),function($cell) use($v){
                      $cell->setBackground("#00FF11");
                    });
                  }
                  else if($v['color']=="B")
                  {
                    $bno++;
                    $sheet->cell($sheet->getCellByColumnAndRow(4*16*3+($v['lines']-1)*3+3+4,54+$v['rows']+ceil($v['group']/5)+15)->getCoordinate(),function($cell) use($v){
                      $cell->setBackground("#002BFF");
                    });
                  }
               }
              else { //group 11.12.13.14

                  if($v['color']=="R")
                  {
                    $rno++;
                    $sheet->cell($sheet->getCellByColumnAndRow((($v['group']%5)-1)*16*3+($v['lines']-1)*3+1+($v['group']%5)-1,54+$v['rows']+ceil($v['group']/5)+15)
                    ->getCoordinate(),function($cell) use($v){
                      $cell->setBackground("#FF0000");
                    });
                  }
                  else if($v['color']=="G")
                  {
                    $gno++;
                    $sheet->cell($sheet->getCellByColumnAndRow((($v['group']%5)-1)*16*3+($v['lines']-1)*3+2+($v['group']%5)-1,54+$v['rows']+ceil($v['group']/5)+15)
                    ->getCoordinate(),function($cell) use($v){
                      $cell->setBackground("#00FF11");
                    });
                  }
                  else if($v['color']=="B")
                  {
                    $bno++;
                    $sheet->cell($sheet->getCellByColumnAndRow((($v['group']%5)-1)*16*3+($v['lines']-1)*3+3+($v['group']%5)-1,54+$v['rows']+ceil($v['group']/5)+15)
                    ->getCoordinate(),function($cell) use($v){
                      $cell->setBackground("#002BFF");
                    });
                  }
                }
              }

            }

          }
          else{
            $sheet->rows($result2);
          }
          $sheet->mergeCells('C3:V3');
          $sheet->getCell('C3')->setValue('Printed Date:');
          $sheet->setBorder('C3','thin');
          $sheet->mergeCells('X3:AS3');
          $sheet->setBorder('X3','thin');
          $sheet->mergeCells('BA3:BN3');
          $sheet->getCell('BA3')->setValue('Lot No:');
          $sheet->setBorder('BA3','thin');
          $sheet->mergeCells('BQ3:CC3');
          $sheet->setBorder('BQ3','thin');
          $sheet->mergeCells('CZ3:DN3');
          $sheet->getCell('CZ3')->setValue('Name:');
          $sheet->setBorder('CZ3','thin');
          $sheet->mergeCells('DR3:EF3');
          $sheet->setBorder('DR3','thin');
          $sheet->mergeCells('CZ9:DN9');
          $sheet->getCell('CZ9')->setValue('Date:');
          $sheet->setBorder('CZ9','thin');
          $sheet->mergeCells('DR9:EF9');
          $sheet->setBorder('DR9','thin');
          $sheet->mergeCells('FK3:GE3');
          $sheet->getCell('FK3')->setValue('Defect record');
          $sheet->setBorder('FK3','thin');
          $sheet->mergeCells('FK6:FO6');
          $sheet->getCell('FK6')->setValue('R:');
          $sheet->setBorder('FK6','thin');
          $sheet->mergeCells('FP6:GE6');
          $sheet->getCell('FP6')->setValue($rno);
          $sheet->setBorder('FP6','thin');
          $sheet->mergeCells('FK7:FO7');
          $sheet->getCell('FK7')->setValue('G:');
          $sheet->setBorder('FK7','thin');
          $sheet->mergeCells('FP7:GE7');
          $sheet->getCell('FP7')->setValue($gno);
          $sheet->setBorder('FP7','thin');
          $sheet->mergeCells('FK8:FO8');
          $sheet->getCell('FK8')->setValue('B:');
          $sheet->setBorder('FK8','thin');
          $sheet->mergeCells('FP8:GE8');
          $sheet->getCell('FP8')->setValue($bno);
          $sheet->setBorder('FP8','thin');
          $sheet->cells('A1:IK15', function($cells) {
            $cells->setFontSize(24);
          });
          $sheet->cells('A1:IK15', function($cells) {
            $cells->setFontSize(24);
          });
          $sheet->cells('A16:IK98',function($cells){
            $cells->setFontSize(24);
          });
        });

        $excel->sheet($title."3", function ($sheet) use ($result2,$chip,$pixel){
          $result3=array();
          foreach($result2 as $v)
          {
            if($v['color']=="R")
            {
              $l=$pixel+$chip;
            }
            else if($v['color']=="G")
            {
              $l=$pixel;
            }
            else if ($v['color']=="B")
            {
              $l=$pixel-$chip;
            }
            if($v['group']<=10)//1.2.3.4.5.6.7.8.9.10
            {
              if($v['group']%5==0)//5.10
              {
                $arr=[(-1)*$pixel*16*4+$v['lines']*(-1)*($pixel)+$pixel,(-1)*$pixel*27*(ceil($v['group']/5)-1)+$v['rows']*(-1)*($pixel)+$l];
                array_push($result3,$arr);
              }
              else //1.2.3.4.6.7.8.9
              {
                $arr=[(-1)*($pixel)*16*($v['group']%5-1)+(-1)*($pixel)*$v['lines']+$pixel,(-1)*$pixel*27*(ceil($v['group']/5)-1)+$v['rows']*(-1)*($pixel)+$l];
                array_push($result3,$arr);
              }
            }
            else //11.12.13.14.15
            {
              if($v['group']%5==0)//15
              {
                $arr=[(-1)*($pixel)*16*4+$v['lines']*(-1)*($pixel)+$pixel,(-1)*$pixel*54+$v['rows']*(-1)*($pixel)+$l];
                array_push($result3,$arr);
              }
              else //11.12.13.14
              {
                $arr=[(-1)*($pixel)*16*($v['group']%5-1)+(-1)*($pixel)*$v['lines']+$pixel,(-1)*$pixel*54+$v['rows']*(-1)*($pixel)+$l];
                array_push($result3,$arr);

              }
            }
          }
          $sheet->rows($result3);
        });


      });

      $ex1->export('xls');

  }
  function median($numbers=array())
  {
    if (!is_array($numbers))
    $numbers = func_get_args();
    rsort($numbers);
    $mid = (count($numbers) / 2);
    return ($mid % 2 != 0) ? $numbers{$mid-1} : (($numbers{$mid-1}) + $numbers{$mid}) / 2;
  }

}
