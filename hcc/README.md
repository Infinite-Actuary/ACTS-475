```bash
ssh username@crane.unl.edu
cd $WORK
lftp -u username@unl.edu,password ftps://ftp.box.com
get data.rdata
exit
sbatch slurm.sh
squeue -u username
```