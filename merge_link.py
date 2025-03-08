from __future__ import print_function
import os
import win32com.client as com

# 設定項目
# ここを必要に応じて書き換えてください
## 入力ファイル名
INPUT_FILENAME = 'input.inpx'
## 出力ファイル名
OUTPUT_FILENAME = 'output.inpx'
# 3.5=道路幅.OSMのデフォルト値.
LANE_WIDTH = 3.5


# 以下、プログラム
if __name__ == '__main__':
    print('Start merge_link')

    # 以下、プログラム
    current_path = os.path.dirname(os.path.abspath(__file__))

    # Vissim起動、シミュレータを回してエッジ情報を生成＆取得
    Vissim = com.gencache.EnsureDispatch("Vissim.Vissim")

    input_filepath = os.path.join(current_path,INPUT_FILENAME)
    Vissim.LoadNet(input_filepath, False)

    Vissim.Net.DynamicAssignment.CreateGraph(0)
    all_edges = Vissim.Net.Edges.GetAll()

    # つなげる対象をエッジ情報から探索
    link_seq_list = []
    for cur_edge in all_edges:
        link_seq = cur_edge.LinkSeq.GetAll()
        if len(link_seq) > 2:
            link_seq_list.append(link_seq)

    # 既存のリンクを削除しながら、リンクのX,Y情報をつなげていく
    for cur_seq in link_seq_list:
        link_list = []
        for cur_link in cur_seq:
            if cur_link.AttValue('IsConn'):
                Vissim.Net.Links.RemoveLink(cur_link)
            else:
                link_list.append(cur_link)

        poly_string = 'LINESTRING('
        poly_list = []
        max_lane_count = 1
        for cur_link in link_list:
            print(f"link_no : {cur_link.AttValue('No')}")
            cur_points = cur_link.LinkPolyPts
            lane_count = len(cur_link.Lanes.GetAll())
            if lane_count > max_lane_count:
                max_lane_count = lane_count
            for cur_pos in cur_points:
                poly_list.append(str(cur_pos.AttValue('X')) + ' ' + str(cur_pos.AttValue('Y')))

            Vissim.Net.Links.RemoveLink(cur_link)

        poly_string += ', '.join(poly_list) + ')'
        Vissim.Net.Links.AddLink(0, poly_string, [LANE_WIDTH] * max_lane_count)


    output_filepath = os.path.join(current_path, OUTPUT_FILENAME)
    Vissim.SaveNetAs(output_filepath)

    Vissim = None

    print('Finish merge_link')