Technicians were task to manually delete log files at remote PC terminals on daily basis. Script was not allowed to run locally at PC terminal, due to rules that script cannot be run unattended.

CMD:del only able to delete network folders that are already been mapped to a local drive.
But numbers of PC terminals counted more then alphabet letters beyond Z.

Decided write some code to ad-hoc map network to drive M:. 
Do deletion, un-map, and map another PC for file folder deletion. 
Run in loop for array of PC.
All PC terminal user login and rights are the same.