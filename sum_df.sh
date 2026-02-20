total_sum=0

for host in <ip.....>; do
   [[ -z "${host// }" ]] && continu;
   used_bytes=$(
    ssh wando@"$host" -p 10022 -o ConnectTimeout=8 -o StrictHostKeyChecking=accept-new \
  "LANG=C df -B1 -x tmpfs -x devtmpfs --output=used,fstype 2>/dev/null \
   | awk 'NR>1 && \$2!=\"nfs\" && \$2!=\"nfs4\" {s+=\$1} END{print s+0}'"
  )
   echo "$host : $used_bytes bytes"
   total_sum=$(( total_sum + used_bytes )); done

echo $total_sum
