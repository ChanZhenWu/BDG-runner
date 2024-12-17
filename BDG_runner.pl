#!/usr/bin/perl
#################################################################
## This script is used to calculate CPK base on 3070 log file.
## Author: noon_chen@apple.com
## V2.5
#################################################################

use Excel::Writer::XLSX;
($sec,$min,$hour,$mday,$mon,$year,$wday,$yday,$isdst) = localtime(time);

#创建一个新的Excel文件
my $log_report = Excel::Writer::XLSX->new('CPK_report'."-".$hour.$min.$sec.'.xlsx');

#添加一个工作表
my $summary = $log_report-> add_worksheet('Summary');
my $workbook = $log_report-> add_worksheet('CPK_report');

$workbook-> freeze_panes(1,11);			# 冻结行、列
$workbook-> set_column(0,0,20);			# 设置列宽
$summary-> set_column(0,5,20);			# 设置列宽
$workbook-> set_row(0,20);				# 设置行高
$summary-> activate();					# 设置初始可见
#$workbook-> protect("drowssap");		# 设置密码
$workbook->set_header('&CUpdated at &D &T');	# 设置页脚
$workbook->set_landscape();				# 设置横排格式
$log_report->set_size(1680, 1180);		# 设置初始窗口尺寸


#新建一个格式
$format_item = $log_report-> add_format(bold=>1, align=>'left', border=>1, size=>12, bg_color=>'cyan');
$format_head = $log_report-> add_format(bold=>1, align=>'vcenter', border=>1, size=>12, bg_color=>'lime');
$format_data = $log_report-> add_format(align=>'center', border=>1);
$format_Fcpk = $log_report-> add_format(align=>'center', border=>1, bg_color=>'orange');
$format_Pcpk = $log_report-> add_format(bold=>0, align=>'center', border=>1, bg_color=>'lime');
$format_Hcpk = $log_report-> add_format(bold=>0, align=>'center', border=>1, bg_color=>'yellow');
$format_FPY  = $log_report-> add_format(align=>'center', border=>1, num_format=> '10');

#写入文件头
$row = 0; $col = 0;
$summary-> write($row, $col, 'SN', $format_head);
$row = 0; $col = 1;
$summary-> write($row, $col, 'Results', $format_head);
$row = 0; $col = 2;
$summary-> write($row, $col, 'TestTime(s)', $format_head);
$row = 0; $col = 3;
$summary-> write($row, $col, 'Criteria', $format_item);
$row = 0; $col = 4;
$summary-> write($row, $col, 'Test Items', $format_item);

$row = 1; $col = 3;
$summary-> write($row, $col, 'CPK >= 1.33', $format_data);
$row = 2; $col = 3;
$summary-> write($row, $col, 'CPK < 1.33', $format_data);
$row = 3; $col = 3;
$summary-> write($row, $col, 'FPY', $format_data);

$row = 1; $col = 4;
$summary-> write($row, $col, '=COUNTIFS(CPK_report!K2:K9999,">=1.33")', $format_data);
$row = 2; $col = 4;
$summary-> write($row, $col, '=COUNTIFS(CPK_report!K2:K9999,"<1.33")', $format_data);
$row = 3; $col = 4;
$summary-> write_formula(3, 4, "=1-(E3/E2)", $format_FPY);  #输出FPY

my $chart = $log_report-> add_chart( type => 'pie', embedded => 1 );
$chart->add_series(
    name       => '=Summary!$B$1',
    categories => '=Summary!$D$2:$D$3',
    values     => '=Summary!$E$2:$E$3',
    data_labels => {value => 1},
);
$summary->insert_chart('D7',$chart,0,0,1.0,1.6);

$row = 0; $col = 0;
$workbook-> write($row, $col, 'Test Items', $format_head);
$row = 0; $col = 1;
$workbook-> write($row, $col, 'TYPE', $format_head);
$row = 0; $col = 2;
$workbook-> write($row, $col, 'Nominal Value', $format_head);
$row = 0; $col = 3;
$workbook-> write($row, $col, 'HiLimit', $format_head);
$row = 0; $col = 4;
$workbook-> write($row, $col, 'LowLimit', $format_head);
$row = 0; $col = 5;
$workbook-> write($row, $col, 'Max (Marginal)', $format_head);
$row = 0; $col = 6;
$workbook-> write($row, $col, 'Min (Marginal)', $format_head);
$row = 0; $col = 7;
$workbook-> write($row, $col, 'Average', $format_head);
$row = 0; $col = 8;
$workbook-> write($row, $col, 'StdDev', $format_head);  #Standard deviation of data
$row = 0; $col = 9;
$workbook-> write($row, $col, 'CP', $format_head);
$row = 0; $col = 10;
$workbook-> write($row, $col, 'CPK', $format_head);


$workbook-> conditional_formatting('J2:K9999',
    {
    	type     => 'cell',
     	criteria => 'between',
     	minimum  => 1.33,
     	maximum  => 10,
     	format   => $format_Pcpk,
    });

$workbook-> conditional_formatting('J2:K9999',
    {
    	type     => 'cell',
     	criteria => 'greater than',
     	value    => 10,
     	format   => $format_Hcpk,
    });

$workbook-> conditional_formatting('J2:K9999',
    {
    	type     => 'cell',
     	criteria => 'greater than',
     	value    => 0,
     	format   => $format_Fcpk,
    });

$workbook-> write_formula(1, 5, "=MAX(L2:AAA2)", $format_data);  		#输出Max
$workbook-> write_formula(1, 6, "=MIN(L2:AAA2)", $format_data);			#输出Min
$workbook-> write_formula(1, 7, "=AVERAGE(L2:AAA2)", $format_data);  	#输出Average
$workbook-> write_formula(1, 8, "=STDEV(L2:AAA2)", $format_data);  		#输出标准差
$workbook-> write_formula(1, 9, "=IF(I2>0,(D2-E2)/6/I2)", $format_data);  #输出CP
$workbook-> write_formula(1, 10, "=MIN((D2-H2),(H2-E2))/I2/3", $format_data);  #输出CPK

######################### create head ####################################################
$row = 1;
$col = 0;
$colSN = 11;
$log_counter = 0;
print "\n","=> extracting header ... ","\n";

@analogfiles = <*.log>;
foreach $analogfiles (@analogfiles)
{
open LogN,"<$analogfiles";
$log_counter++;

	if ($log_counter == 1){
	open NLog,">head";

		$workbook-> write(0, $colSN, $analogfiles, $format_head);	#写入第一个log name
    	$colSN++;

		while($line = <LogN>)
    	{
    	chomp $line;
    	@string = split('\|', $line);
		if ($line =~ "\@BTEST")
    		{
    		#print $string[12]."\n";
    		if ($string[12] eq "1"){$board = "single";}
    		else{$board = "panel";}
    		print $board;
			print "\n".$analogfiles;
    		}

    	elsif ($line =~ "\@BLOCK")
       	{
       		$col = 0;
       		$headN = $string[1];
       		#print "\n".$headN;
       		#$workbook-> write($row, $col, $headN, $format);
       		#$row++;
       		}

        elsif ($line =~ "\@LIM2" and $line =~ "\@A-JUM")    # Jumper
       	{
       		#print $headN, "\r";
       		print NLog $headN, "\r";
       		$workbook-> write($row, $col, $headN, $format_item);					#输出测试名，单项测试
			$workbook-> write($row, 1, substr($line,4,3), $format_data);			#输出TYPE
       		$workbook-> write($row, 3, $string[3], $format_data);					#输出上限值
       		$workbook-> write($row, 4, substr($string[4],0,13), $format_data);		#输出下限值
			$workbook-> write_formula($row, 5, "=MAX(L".($row+1).":AAA".($row+1).")", $format_data);  #输出Max
			$workbook-> write_formula($row, 6, "=MIN(L".($row+1).":AAA".($row+1).")", $format_data);  #输出Min
			$workbook-> write_formula($row, 7, "=AVERAGE(L".($row+1).":AAA".($row+1).")", $format_data);  #输出Average
			$workbook-> write_formula($row, 8, "=STDEV(L".($row+1).":AAA".($row+1).")", $format_data);  #输出标准差
			$workbook-> write_formula($row, 9, "=IF(I".($row+1).">0,(D".($row+1)."-E".($row+1).")/6/I".($row+1).")", $format_data);  #输出CP
			$workbook-> write_formula($row, 10, "=MIN((D".($row+1)."-H".($row+1)."),(H".($row+1)."-E".($row+1)."))/I".($row+1)."/3", $format_data);  #输出CPK

       		$workbook-> conditional_formatting($row, 5,
    			{
    				type     => 'cell',
    			 	criteria => 'greater than',
    			 	value    => "=D".($row+1)."*0.75",
    			 	format   => $format_Fcpk,
    			});
       		$row++;
       		}
        elsif ($line =~ "\@LIM2" and $line =~ "\@A-DIO" and scalar @string == 5)    # Diode
       	{
       		#print $headN, "\r";
       		print NLog $headN, "\r";
       		$workbook-> write($row, $col, $headN, $format_item);					# 输出测试名，单项测试
			$workbook-> write($row, 1, substr($line,4,3), $format_data);			# 输出TYPE
       		$workbook-> write($row, 3, $string[3], $format_data);
       		$workbook-> write($row, 4, substr($string[4],0,13), $format_data);
			$workbook-> write_formula($row, 5, "=MAX(L".($row+1).":AAA".($row+1).")", $format_data);  # 输出Max
			$workbook-> write_formula($row, 6, "=MIN(L".($row+1).":AAA".($row+1).")", $format_data);  # 输出Min
			$workbook-> write_formula($row, 7, "=AVERAGE(L".($row+1).":AAA".($row+1).")", $format_data);  # 输出Average
			$workbook-> write_formula($row, 8, "=STDEV(L".($row+1).":AAA".($row+1).")", $format_data);  # 输出标准差
			$workbook-> write_formula($row, 9, "=IF(I".($row+1).">0,(D".($row+1)."-E".($row+1).")/6/I".($row+1).")", $format_data);  #输出CP
			$workbook-> write_formula($row, 10, "=MIN((D".($row+1)."-H".($row+1)."),(H".($row+1)."-E".($row+1)."))/I".($row+1)."/3", $format_data);  #输出CPK

       		$workbook-> conditional_formatting($row, 5,
    			{
    				type     => 'cell',
    			 	criteria => 'greater than',
    			 	value    => "=D".($row+1)."-(D".($row+1)."-E".($row+1).")*0.2",
    			 	format   => $format_Fcpk,
    			});

       		$workbook-> conditional_formatting($row, 6,
    			{
    				type     => 'cell',
    			 	criteria => 'less than',
    			 	value    => "=E".($row+1)."+(D".($row+1)."-E".($row+1).")*0.2",
    			 	format   => $format_Fcpk,
    			});
    		
       		$row++;
       		}
        elsif ($line =~ "\@LIM2" and $line =~ "\@A-DIO" and scalar @string == 6)    # Diode
       	{
       		$subtitle = substr ($line, 24, rindex($line,"\{@LIM") - 24),"\n";
       		#print "\n".$headN."/".$subtitle; 
       		print NLog $headN."/".$subtitle, "\n";
       		$workbook-> write($row, $col, $headN."/".$subtitle, $format_item);		# 输出测试名，多项测试
			$workbook-> write($row, 1, substr($line,4,3), $format_data);  			# 输出TYPE
       		$workbook-> write($row, 3, $string[4], $format_data);
       		$workbook-> write($row, 4, substr($string[5],0,13), $format_data);
			$workbook-> write_formula($row, 5, "=MAX(L".($row+1).":AAA".($row+1).")", $format_data);  # 输出Max
			$workbook-> write_formula($row, 6, "=MIN(L".($row+1).":AAA".($row+1).")", $format_data);  # 输出Min
			$workbook-> write_formula($row, 7, "=AVERAGE(L".($row+1).":AAA".($row+1).")", $format_data);  # 输出Average
			$workbook-> write_formula($row, 8, "=STDEV(L".($row+1).":AAA".($row+1).")", $format_data);  # 输出标准差
			$workbook-> write_formula($row, 9, "=IF(I".($row+1).">0,(D".($row+1)."-E".($row+1).")/6/I".($row+1).")", $format_data);  #输出CP
			$workbook-> write_formula($row, 10, "=MIN((D".($row+1)."-H".($row+1)."),(H".($row+1)."-E".($row+1)."))/I".($row+1)."/3", $format_data);  #输出CPK

       		$workbook-> conditional_formatting($row, 5,
    			{
    				type     => 'cell',
    			 	criteria => 'greater than',
    			 	value    => "=D".($row+1)."-(D".($row+1)."-E".($row+1).")*0.2",
    			 	format   => $format_Fcpk,
    			});

       		$workbook-> conditional_formatting($row, 6,
    			{
    				type     => 'cell',
    			 	criteria => 'less than',
    			 	value    => "=E".($row+1)."+(D".($row+1)."-E".($row+1).")*0.2",
    			 	format   => $format_Fcpk,
    			});
    		
       		$row++;
       		}
        elsif ($line =~ "\@LIM2" and scalar @string == 5)    # single step Volts
       	{
       		#print $headN, "\r";
       		print NLog $headN, "\r";
       		$workbook-> write($row, $col, $headN, $format_item);					# 输出测试名，单项测试
			$workbook-> write($row, 1, substr($line,4,3), $format_data);			# 输出TYPE
       		$workbook-> write($row, 3, $string[3], $format_data);
       		$workbook-> write($row, 4, substr($string[4],0,13), $format_data);
			$workbook-> write_formula($row, 5, "=MAX(L".($row+1).":AAA".($row+1).")", $format_data);  # 输出Max
			$workbook-> write_formula($row, 6, "=MIN(L".($row+1).":AAA".($row+1).")", $format_data);  # 输出Min
			$workbook-> write_formula($row, 7, "=AVERAGE(L".($row+1).":AAA".($row+1).")", $format_data);  # 输出Average
			$workbook-> write_formula($row, 8, "=STDEV(L".($row+1).":AAA".($row+1).")", $format_data);  # 输出标准差
			$workbook-> write_formula($row, 9, "=IF(I".($row+1).">0,(D".($row+1)."-E".($row+1).")/6/I".($row+1).")", $format_data);  #输出CP
			$workbook-> write_formula($row, 10, "=MIN((D".($row+1)."-H".($row+1)."),(H".($row+1)."-E".($row+1)."))/I".($row+1)."/3", $format_data);  #输出CPK

       		$workbook-> conditional_formatting($row, 5,
    			{
    				type     => 'cell',
    			 	criteria => 'greater than',
    			 	value    => "=D".($row+1)."-(D".($row+1)."-E".($row+1).")*0.2",
    			 	format   => $format_Fcpk,
    			});
       		$workbook-> conditional_formatting($row, 6,
    			{
    				type     => 'cell',
    			 	criteria => 'less than',
    			 	value    => "=E".($row+1)."+(D".($row+1)."-E".($row+1).")*0.2",
    			 	format   => $format_Fcpk,
    			});
    		
       		$row++;
       		}
        elsif ($line =~ "\@LIM3" and scalar @string == 6)     # LCR
       	{
       		#print $headN, "\r";
       		print NLog $headN, "\r";
       		$workbook-> write($row, $col, $headN, $format_item);								# 输出测试名，单项测试
			$workbook-> write($row, 1, substr($line,4,3), $format_data);						# 输出TYPE
       		$workbook-> write($row, 2, substr($line,index($line,"\@LIM")+6,13), $format_data);  # 输出正常值
       		$workbook-> write($row, 3, $string[4], $format_data);
       		$workbook-> write($row, 4, substr($string[5],0,13), $format_data);
			$workbook-> write_formula($row, 5, "=MAX(L".($row+1).":AAA".($row+1).")", $format_data);  # 输出Max
			$workbook-> write_formula($row, 6, "=MIN(L".($row+1).":AAA".($row+1).")", $format_data);  # 输出Min
			$workbook-> write_formula($row, 7, "=AVERAGE(L".($row+1).":AAA".($row+1).")", $format_data);  # 输出Average
			$workbook-> write_formula($row, 8, "=STDEV(L".($row+1).":AAA".($row+1).")", $format_data);  # 输出标准差
			$workbook-> write_formula($row, 9, "=IF(I".($row+1).">0,(D".($row+1)."-E".($row+1).")/6/I".($row+1).")", $format_data);  #输出CP
			$workbook-> write_formula($row, 10, "=MIN((D".($row+1)."-H".($row+1)."),(H".($row+1)."-E".($row+1)."))/I".($row+1)."/3", $format_data);  #输出CPK

			#Hli = 41.25-(41.25-28.05)*0.25
			#Lli = 28.05+(41.25-28.05)*0.25
       		$workbook-> conditional_formatting($row, 5,
    			{
    				type     => 'cell',
    			 	criteria => 'greater than',
    			 	value    => "=D".($row+1)."-(D".($row+1)."-E".($row+1).")*0.2",
    			 	format   => $format_Fcpk,
    			});
       		$workbook-> conditional_formatting($row, 6,
    			{
    				type     => 'cell',
    			 	criteria => 'less than',
    			 	value    => "=E".($row+1)."+(D".($row+1)."-E".($row+1).")*0.2",
    			 	format   => $format_Fcpk,
    			});
    		
       		$row++;
       		}

       	elsif ($line =~ "\@A-")        # Volts
       	{
       		$subtitle = substr ($line, 24, rindex($line,"\{@LIM") - 24),"\n";
       		#print "\n".$headN."/".$subtitle; 
       		print NLog $headN."/".$subtitle, "\n";
       		$workbook-> write($row, $col, $headN."/".$subtitle, $format_item);		# 输出测试名，多项测试
			$workbook-> write($row, 1, substr($line,4,3), $format_data);  			# 输出TYPE
       		$workbook-> write($row, 3, $string[4], $format_data);
       		$workbook-> write($row, 4, substr($string[5],0,13), $format_data);
			$workbook-> write_formula($row, 5, "=MAX(L".($row+1).":AAA".($row+1).")", $format_data);  # 输出Max
			$workbook-> write_formula($row, 6, "=MIN(L".($row+1).":AAA".($row+1).")", $format_data);  # 输出Min
			$workbook-> write_formula($row, 7, "=AVERAGE(L".($row+1).":AAA".($row+1).")", $format_data);  # 输出Average
			$workbook-> write_formula($row, 8, "=STDEV(L".($row+1).":AAA".($row+1).")", $format_data);  # 输出标准差
			$workbook-> write_formula($row, 9, "=IF(I".($row+1).">0,(D".($row+1)."-E".($row+1).")/6/I".($row+1).")", $format_data);  #输出CP
			$workbook-> write_formula($row, 10, "=MIN((D".($row+1)."-H".($row+1)."),(H".($row+1)."-E".($row+1)."))/I".($row+1)."/3", $format_data);  #输出CPK

       		$workbook-> conditional_formatting($row, 5,
    			{
    				type     => 'cell',
    			 	criteria => 'greater than',
    			 	value    => "=D".($row+1)."-(D".($row+1)."-E".($row+1).")*0.2",
    			 	format   => $format_Fcpk,
    			});
       		$workbook-> conditional_formatting($row, 6,
    			{
    				type     => 'cell',
    			 	criteria => 'less than',
    			 	value    => "=E".($row+1)."+(D".($row+1)."-E".($row+1).")*0.2",
    			 	format   => $format_Fcpk,
    			});
			
       		$row++;
       		}
    	}close NLog;}
    elsif($log_counter != 1)
    {
    	print "\n".$analogfiles;
    	#print "\n"." # # # ";
    	$workbook-> write(0, $colSN, $analogfiles, $format_head);		#写入剩余log name
    	$colSN++;
    	}
close LogN;
}

########################## create data ###################################################
$row = 0;
$items = 0;
print "\n","=> extracting data ... ","\n"; 

open (Titles, "<head");
@Titles = <Titles>;
while($title = <@Titles>)	# Items
{
	$col = 11;	# 初始化列数
	$row++;
	chomp $title;
	print "extracting $title\n";
	@analogfiles = <*.log>;
		foreach $analogfiles (@analogfiles)		#log
		{
		$counter = 1;
		open LogN,"<$analogfiles";
			if ($board eq 'single'){
			while($line = <LogN>)	
		    {
		    	chomp $line;
		    	#print substr($line,8,length($line)-11)."\n";
				if ($line =~ "\@BTEST" and $items == 0 and $counter == 1)	# 写SN到Summary中
				{
				#$summary-> write($col-10, 0, substr($line,8,$SNlength), $format_item);
				#if(substr($line,9+$SNlength,2) eq "00"){$summary-> write($col-10, 1, "Pass", $format_Pcpk);}
				#else{$summary-> write($col-10, 1, "Fail", $format_Fcpk);}

				@result = split('\|', $line);  # print "$length[1]\n";  print "$length[2]\n";
				$summary-> write($col-10, 0, $result[1], $format_item);
				if($result[2] eq "00"){$summary-> write($col-10, 1, "Pass", $format_Pcpk);}
				else{$summary-> write($col-10, 1, "Fail", $format_Fcpk);}
				$summary-> write($col-10, 2, $result[4], $format_data);

				#print substr($line,9+$SNlength,2)."\n";
				#print "$items,###,$counter"."\n";
				$counter++;
					}
		    	if (substr($line,8,length($line)-11) eq $title and substr($line,rindex($line,"\|")+1,2) eq "00")		#单项测试数据
					{
						while($line = <LogN>)
						{
							chomp $line;
							if ($line =~ "\@A-")	#result line matching
								{
								$workbook-> write($row, $col, substr ($line,10,13), $format_data); 
								#print substr($line,10,13)."\n";
								#$col++;
								last;
									}
							}
						}
				elsif (substr($line,8,length($line)-11) eq substr($title,0,index($title,"\/")) and substr($line,rindex($line,"\|")+1,2) eq "00")		#多项测试数据
					{
						while($line = <LogN>)
						{
							chomp $line;
							#print substr($title,index($title,"\/")+1,length($title)-index($title,"\/"))."\n";
							#print substr($line,24,index($line,"\@LIM")-25)."\n";
							if ($line =~ "\@A-"	and substr($title,index($title,"\/")+1,length($title)-index($title,"\/")) eq substr($line,24,index($line,"\@LIM")-25) )	#subname matching
								{
								$workbook-> write($row, $col, substr ($line,10,13), $format_data); 
								#print substr($line,10,13)."\n";
								#$col++;
								last;
									}
							if($line eq "\}"){goto Next1;}
							}
						}

				elsif(eof){$col++; last; }

				Next1:}}

			if ($board eq 'panel'){
			while($line = <LogN>)	
		    {
		    	chomp $line;
		    	#print substr($title,index($title,"\%")+1,length($title)-index($title,"\%"))."\n";
		    	#print substr($line,index($line,"\%")+1,length($line)-index($line,"\%")-4)."\n";
				#print "#"."\n";
				if ($line =~ "\@BTEST" and $items == 0 and $counter == 1)	# 写SN到Summary中
				{
				#$summary-> write($col-10, 0, substr($line,8,$SNlength), $format_item);
				#if(substr($line,9+$SNlength,2) eq "00"){$summary-> write($col-10, 1, "Pass", $format_Pcpk);}
				#else{$summary-> write($col-10, 1, "Fail", $format_Fcpk);}

				@result = split('\|', $line);  # print "$length[1]\n";  print "$length[2]\n";
				$summary-> write($col-10, 0, $result[1], $format_item);
				if($result[2] eq "00"){$summary-> write($col-10, 1, "Pass", $format_Pcpk);}
				else{$summary-> write($col-10, 1, "Fail", $format_Fcpk);}
				$summary-> write($col-10, 2, $result[4], $format_data);

				#print substr($line,9+$SNlength,2)."\n";
				#print "$items,###,$counter"."\n";
				$counter++;
					}

				if (substr($line,index($line,"\%")+1,length($line)-index($line,"\%")-4) eq substr($title,index($title,"\%")+1,length($title)-index($title,"\%")) and substr($line,rindex($line,"\|")+1,2) eq "00")		#单项测试数据
					{
					#print substr($line,index($line,"\%")+1,length($line)-index($line,"\%")-4)."\n";
					#print substr($title,index($title,"\%")+1,length($title)-index($title,"\%"))."\n";
						while($line = <LogN>)
						{
							chomp $line;
							if ($line =~ "\@A-")	#result line matching
								{
								#print $line."\n";
								$workbook-> write($row, $col, substr ($line,10,13), $format_data); 
								#print substr($line,10,13)."\n";
								#$col++;
								last;
									}
							}
						}

				elsif ($title =~ "\/" and substr($line,index($line,"\%")+1,length($line)-index($line,"\%")-4) eq substr($title,index($title,"\%")+1,rindex($title,"\/")-length($title)) and substr($line,rindex($line,"\|")+1,2) eq "00")		#多项测试数据
					{
					 #print substr($line,index($line,"\%")+1,length($line)-index($line,"\%")-4)."\n";
					 #print substr($title,index($title,"\%")+1,rindex($title,"\/")-length($title))."\n";
						while($line = <LogN>)
						{
							chomp $line;
							if ($line =~ "\@A-"	and substr($title,index($title,"\/")+1,length($title)-index($title,"\/")) eq substr($line,24,index($line,"\@LIM")-25) )	#subname matching
								{
							#print $line."\n";
								$workbook-> write($row, $col, substr ($line,10,13), $format_data); 
								#print substr($line,10,13)."\n";
								#$col++;
								last;
									}
							if($line eq "\}"){goto Next2;}
							}
						}

				elsif(eof){$col++; last; }

				Next2:}}

		close LogN,"<$analogfiles";
			}
$items++;
	}

close Title;
unlink "head";
$log_report->close();
print "\n	>>> Done.....\n\n";
system 'pause';

exit;


