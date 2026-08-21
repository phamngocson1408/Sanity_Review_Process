export PYTHON311 :=  /Infra/Tools/python/3.11.11/bin/python3.11
export SRAP_HOME := /Infra/Tools/SRAP/srap-compiler
export SANITY_CHECKER :=${SRAP_HOME}/bin/sanity_checker
export PYTHONPATH := ${SRAP_HOME}:${PYTHONPATH}

help:
	-@echo "######################################################################"
	-@echo "make lint		=> Run VC Static on current testcase"
	-@echo "make gui 		=> check lint results in GUI - currently unavailable"
	-@echo " "
	-@echo "make clean		=> Clean up current testcase directory  "
	-@echo " "
	-@echo "make help		=> Generate this help menu      "
	-@echo "######################################################################"

lint_quick:
	-@\mkdir -p logs
	-@\mkdir -p reports
	export QUIT_VC=true;\
	export GUI=false;\
	export SESSION=false;\
	export SAM=false;\
	$(VC_STATIC_HOME)/bin/vc_static_shell   \
		-mode64                             \
	  -lic_wait 5                           \
	  -lic_report                           \
	  -f ${LINT_HOME}/run_lint.tcl \
	  -output_log_file logs/vcst.screen.log
	-@if [ -e ./vcst_rtdb/reports/moresimple.rpt ]; then  \
		if grep -q 'fatal' ./vcst_rtdb/reports/moresimple.rpt; then\
			gvim vcst_session.log; \
		else \
			\cp -r ./vcst_rtdb/reports/moresimple.rpt ./reports/;\
		fi \
	fi

gui:
	export QUIT_VC=true;\
	export GUI=true;\
	export SESSION=false;\
	export SAM=false;\
	$(VC_STATIC_HOME)/bin/vc_static_shell -mode64 -f ${LINT_HOME}/run_lint.tcl

lint:
	-@\mkdir -p logs
	-@\mkdir -p reports
	export QUIT_VC=true;\
	export GUI=false;\
	export SESSION=true;\
	export SAM=false;\
	export SNPSLMD_QUEUE=true; \
	$(VC_STATIC_HOME)/bin/vc_static_shell   \
		-mode64                             \
	  -lic_wait 300                           \
	  -lic_report                           \
	  -f ${LINT_HOME}/run_lint.tcl \
	  -output_log_file logs/vcst.screen.log
	-@if [ -e ./vcst_rtdb/reports/moresimple.rpt ]; then  \
		if grep -q 'fatal' ./vcst_rtdb/reports/moresimple.rpt; then\
			gvim vcst_session.log; \
		else \
			\cp -r ./vcst_rtdb/reports/moresimple.rpt ./reports/;\
		fi \
	fi

lint_post:
	${PYTHON311} ${SANITY_CHECKER} lint reports/report_lint.full.log

vcs_post:
	${PYTHON311} ${SANITY_CHECKER} compile 

excel:
	-@if [ -e ./reports/report_lint.full.log  ]; then  \
		python3 $(CDC_HOME)/script/auto.py \
		reports/report_lint.full.log report_lint.full ;\
	else \
		if [ -e ./reports/report_lint.full.log.gz  ]; then \
			\gunzip reports/* ;\
			python3 $(CDC_HOME)/script/auto.py \
			reports/report_lint.full.log report_lint.full ; \
			\gzip reports/* ;\
		else \
			echo "Any report file is not founded. Rerun LINT."; \
		fi ;\
	fi

waiver:
	-@if [ -e ./vc_waiver.tcl  ]; then  \
		echo "Already waiver.tcl is exist"; \
		echo "Move to another directory original vc_waiver.tcl" ; \
	else \
		echo "copy empty vc_waiver.tcl"; \
		\cp $(CDC_HOME)/vc_waiver.tcl . ;\
		python3 $(CDC_HOME)/script/auto.py \
		report_lint.full.xlsx vc_waiver.tcl 1 2 ;\
	fi


clean:
	-@\rm -rf db_read.log
	-@\rm -rf elab*
	-@\rm -rf upf_command.log 
	-@\rm -rf vcst_command.log
	-@\rm -rf vcst_rtdb
	-@\rm -rf *_rtdb
	-@\rm -rf vcst_rtdb.bak
	-@\rm -rf novas*
	-@\rm -rf reports
	-@\rm -rf *.log
	-@\rm -rf *.log.bak
	-@\rm -rf *.rc
	-@\rm -rf verdiLog
	-@\rm -rf spyglass-*
	-@\rm -rf logs*
	-@\rm -rf session*
	-@\rm -rf work*
	-@\rm -rf vcst_session.log*
	-@\rm -rf .dw_map_file
	-@\rm -rf lint_formal_viol_summary.txt
	-@\rm -rf .vcstPref
	-@\rm -rf .samWrite.lock
	-@\rm -rf result.txt
clean_sam:
	-@\rm -rf sam_*

setup: 
	-@if [ -e ./config_lint.tcl  ]; then  \
		echo "Already config_lint.tcl is exist"; \
		echo "If you want to copy from DM, please delete original config_lint.tcl" ; \
		gvimdiff config_lint.tcl ${LINT_HOME}/config_lint.tcl ; \
	else \
		echo "copy config_lint.tcl"; \
		\cp ${LINT_HOME}/config_lint.tcl ./ ; \
	fi
		
sam:
	-@\mkdir -p logs
	-@\mkdir -p reports
	export QUIT_VC=true;\
	export GUI=false;\
	export SESSION=false;\
	export SAM=true;\
	$(VC_STATIC_HOME)/bin/vc_static_shell -mode64 -f ${LINT_HOME}/run_lint.tcl \
	  -lic_wait 5                           \
	  -lic_report                           \
	  -output_log_file logs/vcst.screen.log
	-@if [ -e ./vcst_rtdb/reports/moresimple.rpt ]; then  \
		if grep -q 'fatal' ./vcst_rtdb/reports/moresimple.rpt; then\
			gvim vcst_session.log; \
		else \
			\cp -r ./vcst_rtdb/reports/moresimple.rpt ./reports/;\
		fi \
	fi


##	-@if [ -e ./vcst_rtdb/reports/moresimple.rpt ]; then  \
##		if grep -q 'fatal' ./vcst_rtdb/reports/moresimple.rpt; then\
##			gvim vcst_session.log; \
##		else \
##			\cp  ./vcst_rtdb/reports/moresimple.rpt ./reports/lint_moresimple_$(DATE)_at$(TIME).rpt; \
##			\mv  ./reports/lint.rpt ./reports/lint_$(DATE)_at$(TIME).rpt; \
##			gvim ./vcst_rtdb/reports/moresimple.rpt; \
##		fi \
##	fi
