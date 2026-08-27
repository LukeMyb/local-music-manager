// 使い方:
// 1. 下記のコード内にある「xxx.xx.xx.xx」をご自身のPC（Tailscale）のIPv4アドレスに書き換えてください。
// 2. 書き換えたコード全体をコピーし、ブラウザのブックマークのURL（アドレス）欄に貼り付けて保存してください。
javascript:(function(){
    fetch('http://xxx.xx.xx.xx:8749/add', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({url: window.location.href})
    })
    .then(response => {
        if(response.ok) alert('Pipelineに送信しました！');
        else alert('送信に失敗しました。サーバーの状態を確認してください。');
    })
    .catch(e => alert('エラーが発生しました（Tailscaleが接続されているか確認してください）: ' + e));
})();