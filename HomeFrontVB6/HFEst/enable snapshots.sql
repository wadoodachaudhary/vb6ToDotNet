--- ENABLE SNAPSHOTS ------------------
/*
insert appoptions(uid,divisionid,optionname,optionvalue) values('',1,'BudgetSnapshotsEnabled','True')
Execute sys.sp_addextendedproperty 'MS_Description','Purchasing:Snap Shots:01:Capture snap shots?','SCHEMA','dbo','TABLE','SecurityGroups','COLUMN','SnapShotAdd'
Execute sys.sp_addextendedproperty 'MS_Description','Purchasing:Snap Shots:02:Delete snap shots?','SCHEMA','dbo','TABLE','SecurityGroups','COLUMN','SnapShotRemove'
*/


--- DISABLE SNAPSHOTS ------------------
/*
delete appoptions where optionname ='budgetsnapshotsenabled'
Execute sys.sp_dropextendedproperty 'MS_Description','SCHEMA','dbo','TABLE','SecurityGroups','COLUMN','SnapShotAdd'
Execute sys.sp_dropextendedproperty 'MS_Description','SCHEMA','dbo','TABLE','SecurityGroups','COLUMN','SnapShotRemove'
*/




