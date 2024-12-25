#az login
#az account set -s "da883fd8-76c8-4ca5-a7e6-8567ae83ce10"
az group create -l uksouth -n rg-hpcuk-temp
az deployment group create `
    --name hpc-cluster-deployment `
    --resource-group rg-hpcuk-temp `
    --template-file ./hpc-cluster-template.json `
    --parameters ./hpc-cluster-parameters.json