package require Tcl 8.4
if { [catch {clipboard clear}] } {
  package require Tk 8.4
  wm withdraw .
}

package provide le_loads 0.1
namespace eval ::le_loads {
	::register::registerMenuActions {}
}


proc write_to_csv {filename data {separator ","}} {
    set fh [open $filename w]
    
    foreach row $data {
        set processed_row [list]
        foreach cell $row {
            # ����������� ��� ������ � ������
            set cell_str [string map {"\"" "\"\""} $cell]
            
            # ���������� ������, ���� �������� �����������
            if {[regexp -- "$separator|\"|\n|\r" $cell_str]} {
                lappend processed_row "\"$cell_str\""
            } else {
                lappend processed_row $cell_str
            }
        }
        puts $fh [join $processed_row $separator]
    }
    close $fh
}}

proc getNets1 { lInstOcc } {
	
	set lNullObj NULL
	set lStatus [DboState]
	#set lNets [list]
	puts "iam here"

	set lPartInst [$lInstOcc GetPartInst $lStatus]
	if {$lPartInst != $lNullObj} {
		set lPinsIter [$lPartInst NewPinsIter $lStatus]
		set lPin [$lPinsIter NextPin $lStatus]
		set lNets [list]
		#lappend lNets "_"		
					
		while {$lPin != $lNullObj} {
			
			
			
			set lPinType [$lPin GetPinType $lStatus]
			
			puts $lPinType
			
			set lNet [$lPin GetNet $lStatus]
			
			set lPinName [DboTclHelper_sMakeCString]
			$lPin GetPinName $lPinName 
			set sPinName [DboTclHelper_sGetConstCharPtr $lPinName]
			puts $sPinName
			
			if {$lPinType == 7} {
				puts 0
				if {$lNet != $lNullObj} {
			
				set lNetName [DboTclHelper_sMakeCString]
				$lNet GetNetName $lNetName 
				set sNetName [DboTclHelper_sGetConstCharPtr $lNetName]
				
				
				puts $sNetName
				
					if {! [string match "*GND*" $sNetName]} {
						
						lappend lNets "$sNetName"
					}

				} else {
					lappend lNets "N/C"
				}
			}
			
			
			set lPin [$lPinsIter NextPin $lStatus]
		}
		delete_DboPartInstPinsIter $lPinsIter
	}
	
	set uniqueNets [lsort -decreasing -unique $lNets]
	
	return $uniqueNets
}


proc getPropertyValue { pDesign pInstOcc pPropertyName } {

  set lNullObj NULL
  set lValue ""
  #$lNullObj

  set lPropName [DboTclHelper_sMakeCString $pPropertyName]
  set lPropValue [DboTclHelper_sMakeCString]

  set lStatus [DboState]
  set lPartInst [$pInstOcc GetPartInst $lStatus]
  
  set lIsVariantInst 0
  if { [$pInstOcc IsVariantPropMapEmpty] == 0} {
    set lIsVariantInst 1
  } elseif { $lPartInst != $lNullObj && [$lPartInst IsVariantPropMapEmpty] == 0} {
    set lIsVariantInst 2
  }


  if {$lIsVariantInst == 1 } {

    set lFindValue [$pInstOcc GetVariantProp $lPropName $lPropValue]

    if { $lFindValue == 1} {
      set lPropValueString [DboTclHelper_sGetConstCharPtr $lPropValue]
      set lDesignCISNotStuffedString [DboTclHelper_sGetConstCharPtr [$pDesign GetCISNotStuffedString]]

      if { $lPropValueString !=  $lDesignCISNotStuffedString} {
        set lValue $lPropValueString
      }
    }

  } elseif {$lIsVariantInst == 2 } {

    set lFindValue [$lPartInst GetVariantProp $lPropName $lPropValue]

    if { $lFindValue == 1} {
      set lPropValueString [DboTclHelper_sGetConstCharPtr $lPropValue]
      set lDesignCISNotStuffedString [DboTclHelper_sGetConstCharPtr [$pDesign GetCISNotStuffedString]]

      if { $lPropValueString !=  $lDesignCISNotStuffedString} {
        set lValue $lPropValueString
        #puts [concat "Variant Part Number (Instance)" $lValue]
      }
    }

  } else {

    set lStatus [$pInstOcc GetEffectivePropStringValue $lPropName $lPropValue]
    if {[$lStatus OK] == 1} {
      set lValue [DboTclHelper_sGetConstCharPtr $lPropValue]
    }
    $lStatus -delete

  }

  return $lValue
}

proc parse_current {current_string} {
    # ������� ������� � ������ � ����� ������
	
	set current_string [string map {" " ""} $current_string]

	
    # ���������� ��������� ��� ���������� ����� � ������� ���������
    if {[regexp {^([-+]?[0-9]*\.?[0-9]+)([a-zA-Z]*)$} $current_string match value unit]} {
		
        # ��������� ������� ��������� � ������ ������� (������ ������)
        set unit_lowercase ""
        foreach char [split $unit ""] {
            append unit_lowercase [string tolower $char]
        }
		
        # ���������� ��������� (������ switch -exact)
        if {$unit_lowercase == "" || $unit_lowercase == "a"} {
            set multiplier 1
        } elseif {$unit_lowercase == "ma"} {
            set multiplier 0.001
        } elseif {$unit_lowercase == "ua"} {
            set multiplier 0.000001
        } elseif {$unit_lowercase == "na"} {
            set multiplier 0.000000001
        } else {
			set multiplier 1
			#puts "����������� ������� ���������: $unit"
            #error "����������� ������� ���������: $unit"
        }
		
        # ���������� ��������� ��� �����
        return [expr {double($value) * $multiplier}]
    } else {
		return ""
		#puts "�������� ������ ����: $current_string"
        #error "�������� ������ ����: $current_string"
    }
}

proc add_or_update_row {table new_row nets} {
    # �������� reference �� ������ ������
    set ref [lindex $new_row 0]

    # ����: ������� ���������� ��� ���
    set found 0

    # �������� �� ���� ������� �������
    for {set i 0} {$i < [llength $table]} {incr i} {
        set row [lindex $table $i]
        set existing_ref [lindex $row 0]

        if {$existing_ref == $ref} {
			set last_idx [expr {[llength $row] - 1}]
            set old_value [lindex $row $last_idx]
            set updated_value [string trim "$old_value $nets"]

            set updated_row [lreplace $row $last_idx $last_idx $updated_value]
            set table [lreplace $table $i $i $updated_row]
            set found 1
            break
        }
    }

    # ���� ���������� ���, ������ ��������� ����� ������
    if {!$found} {
        lappend table $new_row
    }

    # ���������� ���������� �������
    return $table
}


proc getParameters { pDesign pInstOcc } {

	set table [list]

	set lStatus [DboState]
	set lNullObj NULL
	
	set lInstOccIter [$pInstOcc NewChildrenIter $lStatus  $::IterDefs_INSTS]
	$lInstOccIter Sort $lStatus
	set lChildOcc [$lInstOccIter NextOccurrence $lStatus]
	
	set lRawFields {"Reference" "Value" "Supply Current" "VOLTAGE" "CURRENT"}
	
	while { $lChildOcc!= $lNullObj} {
	
		set lId [$lChildOcc GetId $lStatus]
		set lInstOcc [DboOccurrenceToDboInstOccurrence $lChildOcc]
	
		set sRefDes [getPropertyValue $pDesign $lInstOcc "Reference"]
		set sPartType [getPropertyValue $pDesign $lInstOcc "Part Type"]
		set sRefDesNum 0
		if {! [string match "Power ICs*" $sPartType] && ! [string match "Power Supplies*" $sPartType]} {
			if {[string match "D*" $sRefDes]} {
			
				
			
					#set sRefDes_ "[string range $sRefDes 1 end]"
					#set sRefDesNum [expr {$sRefDes_}]
				
					set line [list]

					foreach field $lRawFields {
						if {$field == "Supply Current"} {
							
							set current_string [getPropertyValue $pDesign $lInstOcc $field]
							set current_num [parse_current $current_string]
							lappend line $current_string
							lappend line $current_num
						
						} else {
							lappend line [getPropertyValue $pDesign $lInstOcc $field]
						}
						
						
						
					}
					
					
					
					set Nets [getNets1 $lInstOcc]
					lappend line $Nets
					
					
					set table [add_or_update_row $table $line $Nets]
					
					#lappend table $line
					
					#puts $line
				
			}
		}
		
		set lChildOcc [$lInstOccIter NextOccurrence $lStatus]

	}
	
	delete_DboOccurrenceChildrenIter $lInstOccIter
	$lStatus -delete
	
	return $table
}

proc deleteColumns { data columns_to_delete } {

	# ������� ��� �������� �������� � ������ �������.
	# ������� ������:
	# 	data - ������ �������
	# 	columns_to_delete - ������ �������� ��� ����� �������
	
	set new_data $data

	foreach col [lsort -decreasing $columns_to_delete] {
		set temp_data {}
		foreach row $new_data {
			set new_row [lreplace $row $col $col]
			lappend temp_data $new_row
		}
		set new_data $temp_data
	}
	
	return $new_data
}




proc getActiveDesign {} {
	set lNullObj NULL
	if { [catch {set pDesign [GetActivePMDesign]}] } {
		error "No active design (1)"
    }
	if { $pDesign == $lNullObj } {
		error "No active design (2)"
    }
	return $pDesign
}

proc ::stcCaps::generateLoads { } {
	puts "\nStart"
	
	set pDesign [getActiveDesign]

	set lStatus [DboState]
	set lRootOcc [$pDesign GetRootOccurrence $lStatus]
	
	set table [getParameters $pDesign $lRootOcc]
	
	$lStatus -delete
	
	# ���������
	#set sorted_table [lsort -integer -index end $table]
	#puts "[expr {[llength [lindex $sorted_table 0]] - 1}]"
	#set sorted_table [deleteColumns $sorted_table [expr {[llength [lindex $sorted_table 0]] - 1}]]
	
	# ��������� ���������
	set header {"Reference" "Value" "Supply Current" "SCDigit" "VOLTAGE" "CURRENT" "PowerNets"}
	set sorted_table [linsert $table 0 $header]


	
	
	
	if { [catch {
	
		#set date $::env(DATE)
		#set date [string map [list "." "_"] $::env(DATE)]
		#set time [string map [list ":" "_" "." "_"] $::env(TIME)]
		#set time $::env(TIME)
		
		set currentTime [clock seconds]
		set formattedTime [clock format $currentTime -format "%Y_%m_%d_%H_%M_%S"]
		
		file mkdir "stcReports"
		write_to_csv "stcReports/report_Loads_$formattedTime.csv" $sorted_table
		puts "The report was saved as report_Loads_$formattedTime.csv"
		#exec report_Caps.csv
		if { [catch {
		
			set Excel $::env(EXCELpath)
			exec "$Excel" "stcReports/report_Loads_$formattedTime.csv" &
		
		}] } {
	
			puts "Error : Failed to open file."
		
		}
		
	}] } {
	
		puts "Error : Failed to save file."
		return -1
		
    }
	

	puts "Successful"
	return
}

::stcCaps::registerMenuActions