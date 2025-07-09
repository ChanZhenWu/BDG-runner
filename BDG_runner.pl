#!/usr/bin/perl
#################################################################
## This script is used to calculate CPK base on 3070 log file.
## Author: noon_chen@apple.com
## V4.0
#################################################################

print "\n\tCPK calculator base on i3070 log (v4.0)\n";

use strict;
use warnings;
use Excel::Writer::XLSX;
use Time::HiRes qw(time);
use List::Util qw(min max sum);
use GD::Graph::points;
use PDF::Builder;
use PDF::Table;
mkdir 'Plots';

(my $sec, my $min, my $hour, my $mday, my $mon, my $year,my $wday,my $yday,my $isdst) = localtime(time);
my $start_time = time();

#创建一个新的Excel文件
my $log_report = Excel::Writer::XLSX->new('CPK_report'.'-'.$hour.$min.$sec.'.xlsx');

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
my $format_item = $log_report-> add_format(bold=>1, align=>'left', border=>1, size=>12, bg_color=>'cyan');
my $format_head = $log_report-> add_format(bold=>1, align=>'vcenter', border=>1, size=>12, bg_color=>'lime');
my $format_data = $log_report-> add_format(align=>'center', border=>1);
my $format_Fcpk = $log_report-> add_format(align=>'center', border=>1, bg_color=>'orange');
my $format_Pcpk = $log_report-> add_format(bold=>0, align=>'center', border=>1, bg_color=>'lime');
my $format_Hcpk = $log_report-> add_format(bold=>0, align=>'center', border=>1, bg_color=>'yellow');
my $format_FPY  = $log_report-> add_format(align=>'center', border=>1, num_format=> '10');
my $format_blak = $log_report-> add_format(align=>'center', border=>1, bg_color=>'silver');

#写入文件头
our $row = 0; our $col = 0;
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
$workbook-> write($row, $col, 'Max', $format_head);
$row = 0; $col = 6;
$workbook-> write($row, $col, 'Min', $format_head);
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

$workbook-> conditional_formatting('A1:FFF9999',
{
	type     => 'blanks',
 	format   => $format_blak,
	});

$workbook-> write_formula(1, 5, "=MAX(L2:FFF2)", $format_data);  		#输出Max
$workbook-> write_formula(1, 6, "=MIN(L2:FFF2)", $format_data);			#输出Min
$workbook-> write_formula(1, 7, "=AVERAGE(L2:FFF2)", $format_data);  	#输出Average
$workbook-> write_formula(1, 8, "=STDEV(L2:FFF2)", $format_data);  		#输出标准差
$workbook-> write_formula(1, 9, "=IF(I2>0,(D2-E2)/6/I2)", $format_data);  #输出CP
$workbook-> write_formula(1, 10, "=MIN((D2-H2),(H2-E2))/I2/3", $format_data);  #输出CPK

######################### create head ####################################################
$row = 1;
$col = 0;
my $colSN = 11; my $log_counter = 0;
my $board = ""; my $headN = ""; my $line = ""; my $title = ""; my $subtitle = "";
my @Titles = (); my %DevLim = ();

print "\n=> extracting header ... ","\n";

my @analogfiles = <*.log>;
foreach my $analogfiles (@analogfiles)
{
	open LogN,"<$analogfiles" or warn "\t!!! Failed to open $analogfiles file: $!.\n";
	if ($! eq "No such file or directory"){next;}
	$log_counter++;

	if ($log_counter == 1)
	{
		open NLog,">head";

		$workbook-> write(0, $colSN, $analogfiles, $format_head);	#写入第一个log name
    	$colSN++;

		while(my $line = <LogN>)
    	{
    		chomp $line;
    		my @string = split('\|', $line);

		if ($line =~ "\@BTEST")
    	{
    		#print $string[12]."\n";
    		if ($string[12] eq "1"){$board = "single";}
    		else {$board = "panel";}
    		print $board;
			print "\n".$analogfiles;
    		}

    	elsif ($line =~ "\@BLOCK")
       	{
       		$col = 0;
       		$headN = $string[1];
       		if($board eq "panel"){$headN = substr($string[1], index($string[1],"%")+1);}
       		#print "\n".$headN;
       		#$workbook-> write($row, $col, $headN, $format);
       		#$row++;
       		}

        elsif ($line =~ "\@LIM2" and $line =~ "\@A-JUM")    # Jumper
       	{
       		#print $headN, "\r";
       		print NLog $headN, "\r";
       		push(@Titles, $headN);
       		$workbook-> write($row, $col, $headN, $format_item);					#输出测试名，单项测试
			$workbook-> write($row, 1, substr($line,4,3), $format_data);			#输出TYPE
    		$workbook-> write($row, 2, "-", $format_data);
       		$workbook-> write($row, 3, $string[3], $format_data);					#输出上限值
       		$workbook-> write($row, 4, substr($string[4],0,13), $format_data);		#输出下限值
			$DevLim{$headN} = $string[3].' / '.substr($string[4],0,13).' / '.'-'.' / '.substr($line,4,3);
			$workbook-> write_formula($row, 5, "=MAX(L".($row+1).":FFF".($row+1).")", $format_data);  #输出Max
			$workbook-> write_formula($row, 6, "=MIN(L".($row+1).":FFF".($row+1).")", $format_data);  #输出Min
			$workbook-> write_formula($row, 7, "=AVERAGE(L".($row+1).":FFF".($row+1).")", $format_data);  #输出Average
			$workbook-> write_formula($row, 8, "=STDEV(L".($row+1).":FFF".($row+1).")", $format_data);  #输出标准差
			$workbook-> write_formula($row, 9, "=IF(I".($row+1).">0,(D".($row+1)."-E".($row+1).")/6/I".($row+1).")", $format_data);  #输出CP
			$workbook-> write_formula($row, 10, "=MIN((D".($row+1)."-H".($row+1)."),(H".($row+1)."-E".($row+1)."))/I".($row+1)."/3", $format_data);  #输出CPK
       		$row++;
       		}
        elsif ($line =~ "\@LIM2" and $line =~ "\@A-DIO" and scalar @string == 5)    # Diode
       	{
       		#print $headN, "\r";
       		print NLog $headN, "\r";
       		push(@Titles, $headN);
       		$workbook-> write($row, $col, $headN, $format_item);					# 输出测试名，单项测试
			$workbook-> write($row, 1, substr($line,4,3), $format_data);			# 输出TYPE
    		$workbook-> write($row, 2, "-", $format_data);
       		$workbook-> write($row, 3, $string[3], $format_data);
       		$workbook-> write($row, 4, substr($string[4],0,13), $format_data);
			$DevLim{$headN} = $string[3].' / '.substr($string[4],0,13).' / '.'-'.' / '.substr($line,4,3);
			$workbook-> write_formula($row, 5, "=MAX(L".($row+1).":FFF".($row+1).")", $format_data);  # 输出Max
			$workbook-> write_formula($row, 6, "=MIN(L".($row+1).":FFF".($row+1).")", $format_data);  # 输出Min
			$workbook-> write_formula($row, 7, "=AVERAGE(L".($row+1).":FFF".($row+1).")", $format_data);  # 输出Average
			$workbook-> write_formula($row, 8, "=STDEV(L".($row+1).":FFF".($row+1).")", $format_data);  # 输出标准差
			$workbook-> write_formula($row, 9, "=IF(I".($row+1).">0,(D".($row+1)."-E".($row+1).")/6/I".($row+1).")", $format_data);  #输出CP
			$workbook-> write_formula($row, 10, "=MIN((D".($row+1)."-H".($row+1)."),(H".($row+1)."-E".($row+1)."))/I".($row+1)."/3", $format_data);  #输出CPK
       		$row++;
       		}
        elsif ($line =~ "\@LIM2" and $line =~ "\@A-DIO" and scalar @string == 6)    # Diode
       	{
       		$subtitle = substr ($line, 24, rindex($line,"\{\@LIM") - 24);
       		#print "\n".$headN."/".$subtitle; 
       		print NLog $headN."/".$subtitle, "\n";
       		push(@Titles, $headN."/".$subtitle);
       		$workbook-> write($row, $col, $headN."/".$subtitle, $format_item);		# 输出测试名，多项测试
			$workbook-> write($row, 1, substr($line,4,3), $format_data);  			# 输出TYPE
    		$workbook-> write($row, 2, "-", $format_data);
       		$workbook-> write($row, 3, $string[4], $format_data);
       		$workbook-> write($row, 4, substr($string[5],0,13), $format_data);
			$DevLim{$headN."/".$subtitle} = $string[4].' / '.substr($string[5],0,13).' / '.'-'.' / '.substr($line,4,3);
			$workbook-> write_formula($row, 5, "=MAX(L".($row+1).":FFF".($row+1).")", $format_data);  # 输出Max
			$workbook-> write_formula($row, 6, "=MIN(L".($row+1).":FFF".($row+1).")", $format_data);  # 输出Min
			$workbook-> write_formula($row, 7, "=AVERAGE(L".($row+1).":FFF".($row+1).")", $format_data);  # 输出Average
			$workbook-> write_formula($row, 8, "=STDEV(L".($row+1).":FFF".($row+1).")", $format_data);  # 输出标准差
			$workbook-> write_formula($row, 9, "=IF(I".($row+1).">0,(D".($row+1)."-E".($row+1).")/6/I".($row+1).")", $format_data);  #输出CP
			$workbook-> write_formula($row, 10, "=MIN((D".($row+1)."-H".($row+1)."),(H".($row+1)."-E".($row+1)."))/I".($row+1)."/3", $format_data);  #输出CPK
       		$row++;
       		}
        elsif ($line =~ "\@LIM2" and scalar @string == 5)    # single step Volts
       	{
       		#print $headN, "\r";
       		print NLog $headN, "\r";
       		push(@Titles, $headN);
       		$workbook-> write($row, $col, $headN, $format_item);					# 输出测试名，单项测试
			$workbook-> write($row, 1, substr($line,4,3), $format_data);			# 输出TYPE
    		$workbook-> write($row, 2, "-", $format_data);
       		$workbook-> write($row, 3, $string[3], $format_data);
       		$workbook-> write($row, 4, substr($string[4],0,13), $format_data);
			$DevLim{$headN} = $string[3].' / '.substr($string[4],0,13).' / '.'-'.' / '.substr($line,4,3);
			$workbook-> write_formula($row, 5, "=MAX(L".($row+1).":FFF".($row+1).")", $format_data);  # 输出Max
			$workbook-> write_formula($row, 6, "=MIN(L".($row+1).":FFF".($row+1).")", $format_data);  # 输出Min
			$workbook-> write_formula($row, 7, "=AVERAGE(L".($row+1).":FFF".($row+1).")", $format_data);  # 输出Average
			$workbook-> write_formula($row, 8, "=STDEV(L".($row+1).":FFF".($row+1).")", $format_data);  # 输出标准差
			$workbook-> write_formula($row, 9, "=IF(I".($row+1).">0,(D".($row+1)."-E".($row+1).")/6/I".($row+1).")", $format_data);  #输出CP
			$workbook-> write_formula($row, 10, "=MIN((D".($row+1)."-H".($row+1)."),(H".($row+1)."-E".($row+1)."))/I".($row+1)."/3", $format_data);  #输出CPK
       		$row++;
       		}
        elsif ($line =~ "\@LIM3" and scalar @string == 6)     # LCR
       	{
       		#print $headN, "\r";
       		print NLog $headN, "\r";
       		push(@Titles, $headN);
       		$workbook-> write($row, $col, $headN, $format_item);								# 输出测试名，单项测试
			$workbook-> write($row, 1, substr($line,4,3), $format_data);						# 输出TYPE
       		$workbook-> write($row, 2, substr($line,index($line,"\@LIM")+6,13), $format_data);  # 输出正常值
       		$workbook-> write($row, 3, $string[4], $format_data);
       		$workbook-> write($row, 4, substr($string[5],0,13), $format_data);
			$DevLim{$headN} = $string[4].' / '.substr($string[5],0,13).' / '.substr($line,index($line,"\@LIM")+6,13).' / '.substr($line,4,3);
			$workbook-> write_formula($row, 5, "=MAX(L".($row+1).":FFF".($row+1).")", $format_data);  # 输出Max
			$workbook-> write_formula($row, 6, "=MIN(L".($row+1).":FFF".($row+1).")", $format_data);  # 输出Min
			$workbook-> write_formula($row, 7, "=AVERAGE(L".($row+1).":FFF".($row+1).")", $format_data);  # 输出Average
			$workbook-> write_formula($row, 8, "=STDEV(L".($row+1).":FFF".($row+1).")", $format_data);  # 输出标准差
			$workbook-> write_formula($row, 9, "=IF(I".($row+1).">0,(D".($row+1)."-E".($row+1).")/6/I".($row+1).")", $format_data);  #输出CP
			$workbook-> write_formula($row, 10, "=MIN((D".($row+1)."-H".($row+1)."),(H".($row+1)."-E".($row+1)."))/I".($row+1)."/3", $format_data);  #输出CPK
       		$row++;
       		}

       	elsif ($line =~ "\@A-")        # Volts
       	{
       		$subtitle = substr ($line, 24, rindex($line,"\{\@LIM") - 24);
       		#print "\n".$headN."/".$subtitle; 
       		print NLog $headN."/".$subtitle, "\n";
       		push(@Titles, $headN."/".$subtitle);
       		$workbook-> write($row, $col, $headN."/".$subtitle, $format_item);		# 输出测试名，多项测试
			$workbook-> write($row, 1, substr($line,4,3), $format_data);  			# 输出TYPE
    		$workbook-> write($row, 2, "-", $format_data);
       		$workbook-> write($row, 3, $string[4], $format_data);
       		$workbook-> write($row, 4, substr($string[5],0,13), $format_data);
			$DevLim{$headN."/".$subtitle} = $string[4].' / '.substr($string[5],0,13).' / '.'-'.' / '.substr($line,4,3);
			$workbook-> write_formula($row, 5, "=MAX(L".($row+1).":FFF".($row+1).")", $format_data);  # 输出Max
			$workbook-> write_formula($row, 6, "=MIN(L".($row+1).":FFF".($row+1).")", $format_data);  # 输出Min
			$workbook-> write_formula($row, 7, "=AVERAGE(L".($row+1).":FFF".($row+1).")", $format_data);  # 输出Average
			$workbook-> write_formula($row, 8, "=STDEV(L".($row+1).":FFF".($row+1).")", $format_data);  # 输出标准差
			$workbook-> write_formula($row, 9, "=IF(I".($row+1).">0,(D".($row+1)."-E".($row+1).")/6/I".($row+1).")", $format_data);  #输出CP
			$workbook-> write_formula($row, 10, "=MIN((D".($row+1)."-H".($row+1)."),(H".($row+1)."-E".($row+1)."))/I".($row+1)."/3", $format_data);  #输出CPK
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

print "\n   Scale: ",scalar @Titles,"\n";
# print @Titles,"\n";

# foreach my $key (sort keys %DevLim) {
#     print "$key => $DevLim{$key}\n";
# }

########################## create data ###################################################
print "=> extracting data ... ","\n"; 

our %matrix;
our $value;
our $counter;

%matrix = map { $_ => $Titles[$_] } 0..$#Titles;			# convert array to hash
# my @keys = keys %matrix;
# my $size = @keys;
# print "2 - 哈希大小: $size\n";

foreach my $key (values %matrix) {$matrix{$key} = "";}		# initialize values
# my @keys = keys %matrix;
# my $size = @keys;
# print "2 - 哈希大小: $size\n";

# foreach my $key (keys %matrix) {
#     print $matrix{$key}, "\n";
# }

$row = 0;
$col = 1;
@analogfiles = <*.log>;
foreach my $analogfiles (@analogfiles)		#log
{
	$counter = 1;
	open LogN,"<$analogfiles" or warn "\t!!! Failed to open $analogfiles file: $!.\n";
	if ($! eq "No such file or directory"){next;}
	foreach my $key (keys %matrix) { $matrix{$key} = $matrix{$key}."\t";}		# append default 'tab' for each items.

	while($line = <LogN>)	
    {
    	chomp $line;
    	my @string = split('\|', $line);
		if ($string[0] eq '{@BTEST' and $counter == 1)	# 写SN到Summary中
		{
			$summary-> write($col, 0, $string[1], $format_item);
			if($string[2] eq "00"){$summary-> write($col, 1, "Pass", $format_Pcpk);}
			else{$summary-> write($col, 1, "Fail", $format_Fcpk);}
			$summary-> write($col, 2, $string[4], $format_data);
			$counter++;
			$col++;
			}
    	next unless (substr($line,0,7) eq '{@BLOCK');
		if($string[1] =~ /(^\d+%)/){$string[1] = substr($string[1],index($string[1],"%")+1);}

		my $lines = <LogN>;
		chomp $lines;
		my @lists = split('\|', $lines);

    	if (exists($matrix{$string[1]}))			# 单项测试数据
		{
			if ($lines =~ '@A-')	#result line matching
			{
				#print $lines."\n";
				$value = $matrix{$string[1]};
				#print $value,"\n";
				$value =~ s/\t$//;
				$matrix{$string[1]} = $value.substr ($lines,10,13)."\t";
				}
			}
		elsif (not exists($matrix{$string[1]}) and substr($lists[3],-4) ne 'LIM2' and $lines =~ '@A-')		# 新增单项测试数据
		{
			#print $string[1],"	",substr($lines,2,5),"	",$lines,"\n";
			if (substr($lines,2,5) =~ "(A-RES|A-CAP|A-IND)"){push(@Titles, $string[1]); $DevLim{$string[1]} = $lists[4].' / '.substr($lists[5],0,-2).' / '.$lists[3].' / '.substr($lines,4,3); $matrix{$string[1]} = "\t" x ($col-2).substr ($lines,10,13)."\t";}	#H-L-N
			if (substr($lines,2,5) =~ "(A-JUM|A-DIO)" and scalar @lists eq 5){push(@Titles, $string[1]); $DevLim{$string[1]} = $lists[3].' / '.substr($lists[4],0,-2).' / '.'-'.' / '.substr($lines,4,3); $matrix{$string[1]} = "\t" x ($col-2).substr ($lines,10,13)."\t";}
			}
		elsif (scalar @lists == 6 and exists($matrix{$string[1].'/'.substr($lists[3],0,-6)}))		# 多项测试名
		{
			#print $string[1].'/'.substr($lists[3],0,-6),"	",substr ($lines,10,13),"	",$lines,"\n";
			$value = $matrix{$string[1].'/'.substr($lists[3],0,-6)};
			$value =~ s/\t$//;
			$matrix{$string[1].'/'.substr($lists[3],0,-6)} = $value.substr ($lines,10,13)."\t";
			while($line = <LogN>)
			{
				chomp $line;
				last if ($line !~ '{@A-');
				last if (eof);
				#print $line,"\n";
				my @string1 = split('\|', $line);
				#print $string[1]."/".substr($string1[3],0,length($string1[3])-6),"\n";
				if (scalar @string1 == 6 and exists($matrix{$string[1].'/'.substr($string1[3],0,-6)}))
				{
					$value = $matrix{$string[1].'/'.substr($string1[3],0,-6)};
					$value =~ s/\t$//;
					$matrix{$string[1].'/'.substr($string1[3],0,-6)} = $value.substr ($line,10,13)."\t";
					}
				}
			}
		elsif (scalar @lists == 6 and not exists($matrix{$string[1].'/'.substr($lists[3],0,-6)}))	# 新增多项测试名
		{
			# print $string[1].'/'.substr($lists[3],0,-6),"	",substr($lines,2,5),"	",$lines,"\n";
			# push(@Titles, $string[1].'/'.substr($lists[3],0,-6));
			$DevLim{$string[1].'/'.substr($lists[3],0,-6)} = $lists[4].' / '.substr($lists[5],0,-2).' / '.'-'.' / '.substr($lines,4,3);
			$matrix{$string[1].'/'.substr($lists[3],0,-6)} = "\t" x ($col-2).substr ($lines,10,13)."\t";
			while($line = <LogN>)
			{
				chomp $line;
				last if ($line !~ '{@A-');
				last if (eof);
				#print $line,"\n";
				my @string1 = split('\|', $line);
				#print $string[1]."/".substr($string1[3],0,length($string1[3])-6),"	",$line,"\n";
				if (scalar @string1 == 6 and not exists($matrix{$string[1].'/'.substr($string1[3],0,-6)}))
				{
					push(@Titles, $string[1].'/'.substr($string1[3],0,-6));
					$DevLim{$string[1].'/'.substr($string1[3],0,-6)} = $string1[4].' / '.substr($string1[5],0,-2).' / '.'-'.' / '.substr($line,4,3);
					$matrix{$string[1].'/'.substr($string1[3],0,-6)} = "\t" x ($col-2).substr ($line,10,13)."\t";
					}
				}
			}
		else
		{
			while($line = <LogN>)
			{
				chomp $line;
				last if ($line eq "\}");
				last if (eof);
				my @string1 = split('\|', $line);
				#print "/".substr($string1[3],0,length($string1[3])-6),"\n";
				if ($line =~ '@A-'	and exists($matrix{$string[1]."/".substr($string1[3],0,length($string1[3])-6)}))	#subname matching
				{
					#print $line."\n";
					$value = $matrix{$string[1]."/".substr($string1[3],0,length($string1[3])-6)};
					#print $value,"\n";
					$matrix{$string[1]."/".substr($string1[3],0,length($string1[3])-6)} = $value.substr ($line,10,13)."\t";
					}
				}
			}
		}
	close LogN;
	}

# print "PPDCIN_AON/OUTPUT value is: $matrix{'PPDCIN_AON/OUTPUT'} \n";
# print "r5511 value is: $matrix{'r5511'} \n";

# foreach my $key (sort keys %matrix) {
#     print "$key => $matrix{$key}\n";
# }

# foreach my $key (sort keys %DevLim) {
#     print "$key => $DevLim{$key}\n";
# }


my $data = [
    ['Plot', 'Item', 'Nominal', 'LoLimit', 'Hilimit', 'Minimum', 'Maxmium', 'CPK'],  # 表头
];

foreach my $i (0..@Titles-1)		# create array.
{
	#print $Titles[$i],"\n";
	my @group = split("\t",$matrix{$Titles[$i]});
	@group = map { if (length($_) == 13){$_ = substr($_, 0, 13)} elsif (length($_) > 13){$_ = substr($_, -13)} } @group;		# attain fixed length data (duplicate data).
	my @numeric = grep { $_ =~ /^\S+$/ } @group;		# clear out empty data in array.
	@numeric = map { substr($_, 0, 13) } @numeric;
	my @array = split(" / ",$DevLim{$Titles[$i]});
	my $USL = $array[0];
	my $LSL = $array[1];
	my $Nom = $array[2];
	my $min = min @numeric;
	my $max = max @numeric;
	my $mean = @numeric ? sum(@numeric) / @numeric : warn "!!! array is empty.\n", next;
	my $sigma = sqrt( sum( map { ($_-$mean)**2 } @numeric ) / @numeric );
	if ($sigma == 0){$sigma = 0.1}		# handling sigma = 0.
	my $CPK = min(($USL - $mean),($mean - $LSL))/($sigma*3);
	
	my $data1 = [
    [$i+1, $Titles[$i], $Nom, $LSL, $USL, $min, $max, $CPK],  # 表头
	];
	$data = [@{$data}, @{$data1}];
	}
# my $data = ($head, \@data_array);

# create and config PDF subject.
my $pdf = PDF::Builder->new(-file => 'CPK_report'.'-'.$hour.$min.$sec.'.pdf');
my $page = $pdf->page();
# $page->mediabox('A4');  # 设置 A4 纸张尺寸（595x842点）
$page->mediabox(612, 792);  # 设置 A4 纸张尺寸（595x842点）

my $table = PDF::Table->new();		# output array to PDF.
$table->table(
    $pdf,          # PDF::Builder 对象
    $page,         # 页面对象
    $data,         # 表格数据
    'x'         => 20,		# 左下角起点 X 坐标（距左边距）
    'y'         => 770,		# 左下角起点 Y 坐标（距底边距）
    'w'         => 550,		# 表格总宽度（单位：点）
    'h' 		=> 750,		# 表格总高度（单位：点）
    'next_y'	=> 750,
    'next_h'	=> 700,
    'bg_color_odd'    => "silver",
    # 'bg_color_even'   => "lightblue",
    'border_w'	=> 1,
    'padding'         => 2,	# 单元格内边距
  	# 'padding_right'   => 10,
    font       => $pdf->corefont('Helvetica-Bold'),  # 字体设置
    font_size  => 6,   # 字号
    'justify'	=> 'center',
    header_props => {    # 表头样式
    	'font_size'	=> 8,
        'bg_color'	=> 'green',
        'fg_color'	=> 'orange',
        'justify'	=> 'center',
    },
    cell_props  => [     # 单元格样式（按列索引设置）
        [{ width => 30 }], # 第1列宽度
        [{ width => 30 }],# 第2列宽度
    ]
);
$pdf->save();

# output array to Excel, create diagram and merge into PDF.
$pdf = PDF::Builder->open('CPK_report'.'-'.$hour.$min.$sec.'.pdf');
foreach my $i (0..@Titles-1)
{
	print "analyzing ".$Titles[$i]," ...\n";
	my @group = split("\t",$matrix{$Titles[$i]});
	@group = map { if (length($_) == 13){$_ = substr($_, 0, 13)} elsif (length($_) > 13){$_ = substr($_, -13)} } @group;
	my @numeric = grep { $_ =~ /^\S+$/ } @group;
	@numeric = map { substr($_, 0, 13) } @numeric;
	#print "$Titles[$i] DevLim is: $DevLim{$Titles[$i]} \n";
	my @parametric = split(" / ",$DevLim{$Titles[$i]});
	#print $parametric[2],"\n";
	
	$workbook-> write ($i+1, 3, $parametric[0], $format_data); #Hilimit
	$workbook-> write ($i+1, 4, $parametric[1], $format_data); #Lolimit
	$workbook-> write ($i+1, 2, $parametric[2], $format_data); #Nominal
	$workbook-> write ($i+1, 1, $parametric[3], $format_data); #Type

	$workbook-> write_formula($i+1, 5, "=MAX(L".($i+2).":FFF".($i+2).")", $format_data);  #输出Max
	$workbook-> write_formula($i+1, 6, "=MIN(L".($i+2).":FFF".($i+2).")", $format_data);  #输出Min
	$workbook-> write_formula($i+1, 7, "=AVERAGE(L".($i+2).":FFF".($i+2).")", $format_data);  #输出Average
	$workbook-> write_formula($i+1, 8, "=STDEV(L".($i+2).":FFF".($i+2).")", $format_data);  #输出标准差
	$workbook-> write_formula($i+1, 9, "=IF(I".($i+2).">0,(D".($i+2)."-E".($i+2).")/6/I".($i+2).")", $format_data);  #输出CP
	$workbook-> write_formula($i+1, 10, "=MIN((D".($i+2)."-H".($i+2)."),(H".($i+2)."-E".($i+2)."))/I".($i+2)."/3", $format_data);  #输出CPK

	if ($parametric[3] eq "JUM"){
	$workbook-> conditional_formatting($i+1, 3,
    {
    	type     => 'cell',
     	criteria => 'less than',
     	value    => "=H".($i+2)."+D".($i+2)."*0.5",
     	format   => $format_Fcpk,
    	});
	}
	else{
    $workbook-> conditional_formatting($i+1, 3,
    {
    	type     => 'cell',
     	criteria => 'less than',
     	value    => "=H".($i+2)."+(D".($i+2)."-E".($i+2).")*0.25",
     	format   => $format_Fcpk,
    	});
    $workbook-> conditional_formatting($i+1, 4,
    {
    	type     => 'cell',
     	criteria => 'greater than',
     	value    => "=H".($i+2)."-(D".($i+2)."-E".($i+2).")*0.25",
     	format   => $format_Fcpk,
    	});
	}
	
	$workbook-> write_row ($i+1, 11, \@group, $format_data);		# output array to Excel.
	next if (scalar(@group) == 0);

	# print $DevLim{$Titles[$i]},"\n";
	my @array = split(" / ",$DevLim{$Titles[$i]});
	my $USL = $array[0];
	my $LSL = $array[1];
	my $Nom = $array[2];

	# print $USL,"\n";
	# print $LSL,"\n";
	# print $Nom,"\n";
	if ($USL eq "+9.999999E+99"){$USL = (max @numeric)*1.5;}
	if ($Nom eq "-"){$Nom = $USL*0.9;}

	my @x; my @y; my $y_max; my $y_min;
	my @LoLi; my @HiLi; my @Nomi;
	foreach my $s (0..@numeric-1)
	{
		push @x, '';
		push @y, $numeric[$s];
		push @LoLi, $LSL;
		push @HiLi, $USL;
		push @Nomi, $Nom;
		}

	my $min = min @numeric;
	my $max = max @numeric;
	my $mean = @numeric ? sum(@numeric) / @numeric : warn "!!! array is empty.\n";
	print "	Min: $min\n";
	print "	Max: $max\n";
	print "	Ave: $mean\n"; 

	# calculate StdDev
    my $sigma = sqrt( sum( map { ($_-$mean)**2 } @numeric ) / @numeric );
	if ($sigma == 0){$sigma = 0.1}		# handling sigma = 0.
	printf "	StdDev: %.4f\n", $sigma;
	
	# calculate CPK
    my $CPK = min(($USL - $mean),($mean - $LSL))/($sigma*3);
	printf "	CPK: %.4f\n", $CPK;
	$CPK = sprintf("%.4f", $CPK);

	if ($max > 0 and $max < $USL){$y_max = $USL * 1.1}
	if ($max < 0 and $max > $USL){$y_max = $max * 0.9}
	if ($min > 0 and $min > $LSL){$y_min = $LSL * 0.9}
	if ($min < 0 and $min < $LSL){$y_min = $min * 1.1}
	if ($LSL < 0){$y_min = $LSL * 1.1}
	if ($max > 0 and $USL > $max*2){$y_max = $max * 1.3}

	# PNG visualize object configuration
	my $graph = GD::Graph::points->new(700, 500);
	$graph->set(
	    title			=> uc($Titles[$i])." data distribution",
	    x_label			=> "Count = ".scalar @numeric.",   Min = $min,   Max = $max,   CPK = $CPK",
	    y_label			=> 'Tolerance: (H)'.$USL.' / (L)'.$LSL,
	    markers			=> [7, 3, 9, 8],
	    dclrs			=> ['marine', 'lred', 'lred', 'green'],
		transparent		=> 0,
	    legend_placement => 'RC',
	    marker_size		=> 5,
		# x_label_skip	=> 100,
    	y_tick_number	=> 10,            	# Y 轴刻度数量
    	y_max_value		=> $y_max,			# Y 轴最大值
    	y_min_value		=> $y_min,			# Y 轴最小值
    	#y_tick_length	=> 10,            	# Y 轴刻标长度
    	y_long_ticks    => 1,				# Y 轴长刻度
    	x_tick_length	=> 10,            	# X 轴刻标长度
		axis_space      => 10,				# 轴线到文字的距离
		x_label_position	=> 1/2,			# X 轴承标位置
		# bgclr			=> 'gray',
	);

	# format array（X-Y coordinate pair）
	my @data = (\@x, \@y, \@HiLi, \@LoLi, \@Nomi);

	$graph->set_legend('Measured Value', 'High Limit', 'Low Limit', 'Nominal Value');

	my $PNGTitle = $Titles[$i];
	if ($PNGTitle =~ "\/"){$PNGTitle =~ s/\//\%/i;}

	# generate PNG diagram.
	open my $fh, '>', 'Plots/'.uc($PNGTitle).'.png' or die $!;
	binmode $fh;
	print $fh $graph->plot(\@data)->png;
	close $fh;

	# create hyperlink into excel.
	$workbook-> write_url ($i+1, 0, 'Plots/'.$PNGTitle.'.png', $Titles[$i], $format_item);

	# 添加A4页面（尺寸为595x842点）
	my $page = $pdf->page();
	$page->mediabox(750, 550);

	# 加载图片（支持PNG/JPEG格式）
	my $image = $pdf->image('Plots/'.$PNGTitle.'.png');  
	$page->object($image, 25, 25);	# 指定坐标位置

	# 插入单行文本
	my $font = $pdf->corefont("Helvetica");
	my $text = $page->text();
	$text->fillcolor('white');
	$text->font($font, 12);
	$text->translate(5, 540);
	$text->text($Titles[$i]);

	}

unlink "head";
$log_report->close();

# 保存PDF文件
$pdf->save('CPK_report'.'-'.$hour.$min.$sec.'.pdf');

my $end_time = time();
my $duration = $end_time - $start_time;
printf "\n	runtime: %.4f Sec\n", $duration;

print "\n	>>> Done .....\n\n";
# <STDIN>;


