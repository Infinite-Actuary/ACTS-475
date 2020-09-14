#!/bin/bash
#SBATCH --ntasks=8
#SBATCH --nodes=1
#SBATCH --time=04:00:00
#SBATCH --mem-per-cpu=2048
#SBATCH --workdir=/work/actuarialsci/ccorder3/acts875
#SBATCH --job-name=rf-training
module load R/3.5
R CMD BATCH train_models.r