# PICT-PAPP
Pairwise Independent Combinatorial Tool - Pre-And-Post Processor

EXCEL上で動作するPairwise組み合わせテスト設計を支援するツールPICT-PAPP(Pairwise Independent Combinatorial Tool - Pre-And-Post Processor)です。

Pairwiseの組み合わせ生成そのものは、よく知られた[CIT-BACH](http://www-ise4.ist.osaka-u.ac.jp/~t-tutiya/CIT/)または[PICT](https://github.com/Microsoft/pict)にやってもらうものです。
類似のツールとしては、これまたよく知られた [PictMaster](https://ja.osdn.net/projects/pictmaster/) があります。どちらもEXCELのVBAで実装されています。外部ツールCIT-BACHとPICTへの入力ファイルを生成し、指定した外部ツールであるCIT-BACHもしくはPICTを実行し、結果を取り込みます。
（「既存ツールがあるのに、なんでまた作るかな？」と思われた方には、ぜひ[詳しい説明をQiitaの方で書きました](https://qiita.com/drafts/a46bb9bb5aca490f90f5)ので、補足説明まで読んでいただけると嬉しいです。補足説明の方が長かったりしますが、、、）

同時にこのツールを使って組み合わせテストを設計するための手順書もアップします。

どんな特徴があるのかすぐにわかるように、サンプルデータと、実行済みのデータを残したファイルもここに置いておきました。どんな結果が得られるのか確認してみてください。
