<?php
/**
* RoomUtilizationPlugin.class.php
*
* 
*
*
* @author		André Noack <noack@data-quest.de>, Suchi & Berg GmbH <info@data-quest.de>
* @version		$Id:$
*/
// +---------------------------------------------------------------------------+
// This file is part of Stud.IP
// RoomUtilizationPlugin.class.php
// stellt eine Übersicht der Auslastung von Raumgruppen pro Semester dar
// Copyright (C) 2007 André Noack <noack@data-quest.de>
// Suchi & Berg GmbH <info@data-quest.de>
// +---------------------------------------------------------------------------+
// This program is free software; you can redistribute it and/or
// modify it under the terms of the GNU General Public License
// as published by the Free Software Foundation; either version 2
// of the License, or any later version.
// +---------------------------------------------------------------------------+
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// You should have received a copy of the GNU General Public License
// along with this program; if not, write to the Free Software
// Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
// +---------------------------------------------------------------------------+

require_once "lib/resources/lib/RoomGroups.class.php";
require_once "lib/resources/lib/AssignEventList.class.php";


class RoomUtilizationPlugin extends AbstractStudIPAdministrationPlugin {
	
	var $group_id = 0;
	
	function RoomUtilizationPlugin() {
		
		parent::AbstractStudIPAdministrationPlugin();
		
		$this->setPluginiconname("img/plugin.png");
		
		$navigation =& new PluginNavigation();
		$navigation->setDisplayname($this->getDisplayname());
		$this->setNavigation($navigation);
		$top_navigation =& new PluginNavigation();
		$top_navigation->setDisplayname($this->getDisplayname());
		$this->setTopnavigation($top_navigation);
	}
	
	function getDisplayname() {
		return _("Raumauslastung");
	}
	
	
	function actionShow() {
		if(isset($_REQUEST['sem_schedule_choose'])) $GLOBALS['_default_sem'] = $_REQUEST['sem_schedule_choose'];
		if(isset($_REQUEST['group_schedule_choose_group'])) $this->group_id = $_REQUEST['group_schedule_choose_group'];
		$sem =& new SemesterData();
		$semester = $sem->getSemesterData($GLOBALS['_default_sem']);
		if(!$semester) {
			$semester = $sem->getCurrentSemesterData();
			$GLOBALS['_default_sem'] = $semester['semester_id'];
		}
		$room_groups = RoomGroups::GetInstance();
		$group = $room_groups->getGroupContent($this->group_id);
		foreach($group as $resource_id){
			$assign_events = new AssignEventList($semester['beginn'], $semester['ende'], $resource_id, '', '', TRUE);
			$data[$resource_id]['name'] = getResourceObjectName($resource_id);
			while ($event = $assign_events->nextEvent()) {
				if( ($duration = $event->getEnd() - $event->getBegin()) > 0){
					$data[$resource_id]['utilization'][date('W', $event->getBegin())] += $duration;
				}
			}
		}
		$data['group']['name'] = $room_groups->getGroupName($this->group_id);
		foreach($data as $key => $values){
			if(is_array($values['utilization'])){
				foreach($values['utilization'] as $kw => $util){
					$data['group']['utilization'][$kw] += $util;
					$data[$key]['utilization_sum'] += $util;
				}
			} else {
				$data[$key]['utilization_sum'] = 0;
			}
		}
		$data['group']['utilization_sum'] = is_array($data['group']['utilization']) ? array_sum($data['group']['utilization']) : 0;
		if($_REQUEST['send_as_xls_x']){
			$tmpfile = basename($this->createResultXLS($data,$semester));
			if($tmpfile){
				ob_end_clean();
				header('Location: ' . getDownloadLink( $tmpfile, _("raumauslastung.xls"), 4));
				page_close();
				die;
			}
		}
		$this->showNavigator();
		$this->showResult($data, $semester);
	}
	
	function getKWArray($start, $end){
		$ret = array();
		for($i = $start; $i <= $end; $i = strtotime('+1 week', $i)){
			$ret[] = date('W', $i);
		}
		if(date('W', $end) != $ret[count($ret)-1]) $ret[] = date('W', $end);
		return $ret;
	}
	
	function showResult($data,$semester){
		global $cssSw;
		$kw_sem = $this->getKWArray($semester['beginn'], $semester['ende']);
		$kw_vorles = $this->getKWArray($semester['vorles_beginn'], $semester['vorles_ende']);
		$count_kw_sem = count($kw_sem);
		$count_kw_vorles = count($kw_vorles);
		$cssSw->switchClass();
		?>
		<div class="<? echo $cssSw->getClass() ?>" align="center" style="margin:5px;">
		<b>
		<? printf(_("Anzeige des Semesters: %s"), htmlReady($semester['name']));
		echo '<br>' . date ("d.m.Y", $semester['beginn']), " - ", date ("d.m.Y", $semester['ende']);
		?>
		</b>
		<br />
		</div>
		<table border="0" celpadding="2" cellspacing="2" width="99%" align="center">
		<tr>
		<th width="10%" style="font-size:80%"><?=_("Name")?></th>
		<th width="10%" style="font-size:80%"><?=_("Gesamt")?></th>
		<th width="10%" style="font-size:80%"><?=_("ø Sem")?></th>
		<th width="10%" style="font-size:80%"><?=_("ø Vorles")?></th>
		<?foreach($kw_sem as $kw) echo '<th style="font-size:80%">KW'.$kw.'</th>';?>
		</tr>
		<?
		foreach($data as $resource_id => $values){
			$cssSw->switchClass();
			if($resource_id == 'group'){
				echo chr(10) . '<tr><td  ' . $cssSw->getFullClass() . ' colspan="'.($count_kw_sem + 4).'" style="font-size:80%">---</td></tr>';
			}
			echo chr(10) . '<tr>';
			echo chr(10) . '<td ' . $cssSw->getFullClass() . ' style="font-size:80%;text-align:center;">'.($resource_id != 'group' ? '<a href="'.UrlHelper::getLink('resources.php?view=view_sem_schedule&navigate=1&sem_schedule_choose=0&show_object='.$resource_id).'">' : '') . htmlReady($values['name']) . ($resource_id != 'group' ? '</a>':'').'</td>';
			echo chr(10) . '<td ' . $cssSw->getFullClass() . ' style="font-size:80%;text-align:center;">' . htmlReady(round($values['utilization_sum']/60/60,2)) . '</td>';
			echo chr(10) . '<td ' . $cssSw->getFullClass() . ' style="font-size:80%;text-align:center;">' . htmlReady(round($values['utilization_sum']/$count_kw_sem/60/60,1)) . '</td>';
			echo chr(10) . '<td ' . $cssSw->getFullClass() . ' style="font-size:80%;text-align:center;">' . htmlReady(round($values['utilization_sum']/$count_kw_vorles/60/60,1)) . '</td>';
			foreach($kw_sem as $kw) echo chr(10) . '<td ' . $cssSw->getFullClass() . ' style="font-size:80%;text-align:center;">'.htmlReady(round($values['utilization'][$kw]/60/60,2)).'</td>';
			echo chr(10) . '</tr>';
		}
	}
	
	function createResultXLS($data,$semester){
		require_once "vendor/write_excel/OLEwriter.php";
		require_once "vendor/write_excel/BIFFwriter.php";
		require_once "vendor/write_excel/Worksheet.php";
		require_once "vendor/write_excel/Workbook.php";
		
		global $TMP_PATH;
		$kw_sem = $this->getKWArray($semester['beginn'], $semester['ende']);
		$kw_vorles = $this->getKWArray($semester['vorles_beginn'], $semester['vorles_ende']);
		$count_kw_sem = count($kw_sem);
		$count_kw_vorles = count($kw_vorles);
		$tmpfile = $TMP_PATH . '/' . md5(uniqid('write_excel',1));
		// Creating a workbook
		$workbook = new Workbook($tmpfile);
		$head_format =& $workbook->addformat();
		$head_format->set_size(12);
		$head_format->set_bold();
		$head_format->set_align("left");
		$head_format->set_align("vcenter");
		
		$head_format_merged =& $workbook->addformat();
		$head_format_merged->set_size(12);
		$head_format_merged->set_bold();
		$head_format_merged->set_align("left");
		$head_format_merged->set_align("vcenter");
		$head_format_merged->set_merge();
		$head_format_merged->set_text_wrap();
		
		$caption_format =& $workbook->addformat();
		$caption_format->set_size(10);
		$caption_format->set_align("left");
		$caption_format->set_align("vcenter");
		$caption_format->set_bold();
		
		$data_format =& $workbook->addformat();
		$data_format->set_size(10);
		$data_format->set_align("left");
		$data_format->set_align("vcenter");

		$caption_format_merged =& $workbook->addformat();
		$caption_format_merged->set_size(10);
		$caption_format_merged->set_merge();
		$caption_format_merged->set_align("left");
		$caption_format_merged->set_align("vcenter");
		$caption_format_merged->set_bold();


		// Creating the first worksheet
		$worksheet1 =& $workbook->addworksheet(_("Raumauslastung"));
			$worksheet1->set_row(0, 20);
			$worksheet1->write_string(0, 0, _("Raumauslastung") . ' - ' . $data['group']['name'] ,$head_format);
			$worksheet1->set_row(1, 20);
			$worksheet1->write_string(1, 0, sprintf(_("Anzeige des Semesters: %s; %s"), $semester['name'],date ("d.m.Y", $semester['beginn']) . " - " . date ("d.m.Y", $semester['ende'])), $head_format);
			
			for($i = 1; $i < ($count_kw_sem + 4); ++$i){
				$worksheet1->write_blank(0,$i,$head_format);
				$worksheet1->write_blank(1,$i,$head_format);
			}
			
			$row = 2;
			++$row;
			$worksheet1->write_string($row, 0 , _("Name"), $caption_format);
			$worksheet1->set_column(0, 0, 50);
			$worksheet1->write_string($row, 1 , _("Gesamt"), $caption_format);
			$worksheet1->set_column(0, 1, 15);
			$worksheet1->write_string($row, 2 , _("ø Sem"), $caption_format);
			$worksheet1->set_column(0, 2, 15);
			$worksheet1->write_string($row, 3 , _("ø Vorles"), $caption_format);
			$worksheet1->set_column(0, 3, 15);
			$c = 3;
			foreach($kw_sem as $kw) {
				$worksheet1->write_string($row, ++$c, 'KW'.$kw , $caption_format);
				$worksheet1->set_column(0, $c, 10);
			}
			++$row;
			foreach($data as $resource_id => $values){
				$worksheet1->write_string($row, 0, $values['name'], $data_format);
				$worksheet1->write_number($row, 1, round($values['utilization_sum']/60/60,2), $data_format);
				$worksheet1->write_number($row, 2, round($values['utilization_sum']/$count_kw_sem/60/60,1), $data_format);
				$worksheet1->write_number($row, 3, round($values['utilization_sum']/$count_kw_vorles/60/60,1), $data_format);
				$c = 3;
				foreach($kw_sem as $kw) {
					$worksheet1->write_number($row, ++$c, round($values['utilization'][$kw]/60/60,2), $data_format);
				}
				++$row;
			}
			$workbook->close();
		return $tmpfile;
	}
	
	function showNavigator() {
		global $cssSw, $PHP_SELF;
		$cssSw = new cssClassSwitcher();

			?>
			<table border="0" celpadding="2" cellspacing="0">
			<form method="POST" name="schedule_form" action="<?echo $_SERVER["REQUEST_URI"] ?>">
			<tr>
			<td class="<? echo $cssSw->getClass() ?>" width="96%" colspan="4">
			<font size="-1">&nbsp;</font></td>
			</tr>
			<tr>
			<td class="<? echo $cssSw->getClass() ?>" width="4%" rowspan="2">&nbsp;
			</td>
			<td class="<? echo $cssSw->getClass() ?>" width="40%" valign="top">
			<font size=-1><b><?=_("Semester:")?></b></font>
			
			</td>
			<td class="<? echo $cssSw->getClass() ?>"valign="top">
			<font size="-1">
			<?=_("Eine Raumgruppe ausw&auml;hlen")?>
			</font>
			</td>
			<td class="<? echo $cssSw->getClass() ?>"><font size="-1">
			&nbsp;</font>
			</td>
			</tr>
			<tr>
			<td class="<? echo $cssSw->getClass() ?>" width="40%" valign="middle">
			<font size="-1">
			<?=SemesterData::GetSemesterSelector(array('name' => 'sem_schedule_choose', 'onChange' => 'document.schedule_form.submit()'), $GLOBALS['_default_sem'],'semester_id',false)?>
			<input type="IMAGE" name="jump" align="absbottom" border="0"<? echo makeButton("auswaehlen", "src") ?> /><br />
			</font>
			</td>
			<td class="<? echo $cssSw->getClass() ?>" valign="middle"><font size="-1">
			<select name="group_schedule_choose_group" onChange="document.schedule_form.submit()">
			<?
			$room_group = RoomGroups::GetInstance();
			foreach($room_group->getAvailableGroups() as $gid){
				echo '<option value="'.$gid.'" '
				. ($this->group_id == $gid ? 'selected' : '') . '>'
				.htmlReady(my_substr($room_group->getGroupName($gid),0,85))
				.' ('.$room_group->getGroupCount($gid).')</option>';
			}
			?>
			</select>
			</font>
			</td>
			<td class="<? echo $cssSw->getClass() ?>" valign="middle" align="left"><font size="-1">
			<input type="IMAGE" name="group_schedule_start" align="middle" <?=makeButton("auswaehlen", "src") ?> border=0  />
			</font>
			</td>
			</tr>
			<tr>
			<td class="<? echo $cssSw->getClass() ?>" colspan="4" align="center"><font size="-1">
			<input type="IMAGE" name="send_as_xls" align="middle" <?=makeButton("herunterladen", "src") ?> border=0  />
			</font>
			</td>
			</tr>
			</table>
			</form>
			<?
	}
}
?>
