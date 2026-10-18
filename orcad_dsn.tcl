# Dump every placed part of an OrCAD .DSN to a UTF-8 TSV file.
# usage: tclsh orcad_dsn.tcl <dbo_dll> <design.DSN> <out.tsv>
# out lines: P<TAB>Reference<TAB>Value<TAB>PCB Footprint   (one per placed section)
#            S<TAB>schematic name
# exit code: 0 ok / 2 cannot open design
lassign $argv dll dsn out
load $dll DboTclWriteBasic

set session [DboTclHelper_sCreateSession]
set status [DboState]
set design [$session GetDesignAndSchematics [DboTclHelper_sMakeCString $dsn] $status]
if {$design == "NULL" || [$status Code] != 0} {
    puts stderr "open failed: code=[$status Code]"
    exit 2
}

set f [open $out w]
fconfigure $f -encoding utf-8 -translation lf

proc prop {inst name} {
    set v [DboTclHelper_sMakeCString]
    $inst GetEffectivePropStringValue [DboTclHelper_sMakeCString $name] $v
    return [string map {"\t" " " "\n" " " "\r" ""} [DboTclHelper_sGetConstCharPtr $v]]
}

set vit [$design NewViewsIter $status $::IterDefs_SCHEMATICS]
set view [$vit NextView $status]
while {$view != "NULL"} {
    set sch [DboViewToDboSchematic $view]
    set nm [DboTclHelper_sMakeCString]
    $sch GetName $nm
    puts $f "S\t[DboTclHelper_sGetConstCharPtr $nm]"
    set pit [$sch NewPagesIter $status]
    set page [$pit NextPage $status]
    while {$page != "NULL"} {
        set iit [$page NewPartInstsIter $status]
        set inst [$iit NextPartInst $status]
        while {$inst != "NULL"} {
            puts $f "P\t[prop $inst {Reference}]\t[prop $inst {Value}]\t[prop $inst {PCB Footprint}]"
            set inst [$iit NextPartInst $status]
        }
        delete_DboPagePartInstsIter $iit
        set page [$pit NextPage $status]
    }
    set view [$vit NextView $status]
}
close $f
exit 0
