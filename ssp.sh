#!/bin/bash

#winres="$(xrandr | grep 'current')" #save result to winres
#IFS=' ' read -r -a array <<< $winres #split and save to array

#get current resolution
#for index in "${!array[@]}"
#do
    #if [ ${array[index]} == 'current' ]
    #then
        #x=${array[index+1]}
        #y=${array[index+3]::-1}
        #break
    #fi
#done

#run store item search
konsole -geometry +0-0 -p 'LocalTabTitleFormat=Sari-Sari Store Price List' -p 'TerminalColumns=59' -p 'TerminalRows=63' -e /home/b/projects/sari_store_prices/ssp.py &