# esxi-batch-clone
esxi-batch-clone.ps1，读取 excel 文件配置, 通过调用 vmware-powercli api 操作 vcenter 实现引用“虚拟主机模板”快速批量创建并配置主机（例如：主机名，网络配置，网络标签）

# install
VMware-PowerCLI-5.5.0-1931983.exe

# create
esxi vm template

# edit
host_list.xlsx

# run clone
use powercli run esxi-batch-clone.ps1
