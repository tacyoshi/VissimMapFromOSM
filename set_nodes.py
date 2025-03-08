from __future__ import print_function
import os
import win32com.client as com
import math


# 設定項目
# ここを必要に応じて書き換えてください
## 入力ファイル名
INPUT_FILENAME = 'input.inpx'
## 出力ファイル名
OUTPUT_FILENAME = 'output.inpx'
## 基準の道路幅
LANE_WIDTH = 3.5
## 探索範囲
SEARCH_DISTANCE = 15
## 範囲のマージン
RECT_MARGIN = 20 #(RECT_MARGEN > 0)


# 以下、プログラム

# 指定地点が範囲内に入っているかどうか判定
def is_point_inside_rectangle(point, rectangle):
    x, y = point
    x1, y1 = rectangle[0]
    x2, y2 = rectangle[1]

    if x1 <= x <= x2 and y1 <= y <= y2:
        return True
    else:
        return False

# 複数の指定地点から範囲を計算する
def get_rectangle_points(point_info_list,margin=0):

    x_points = []
    y_points = []
    for point_info in point_info_list:
        x_points.append(links[point_info['key']][point_info['target']][0])
        y_points.append(links[point_info['key']][point_info['target']][1])
    
    return [[min(x_points)-margin,min(y_points)-margin],[min(x_points)-margin,max(y_points)+margin],[max(x_points)+margin,max(y_points)+margin],[max(x_points)+margin,min(y_points)-margin]]


# ２点間の距離を計算
def calculate_distance(point1,point2):
    return math.dist(point1,point2)


# ノードにいれるリンクの端点を計算
def select_points(base_point_info_list):

    point_info_list = base_point_info_list.copy()


    center_point = None
    is_added = False

    # ノードの範囲を計算＋重点を計算
    rectangle_points =  get_rectangle_points(point_info_list)
    center_point = [(rectangle_points[0][0] + rectangle_points[1][0])/2,(rectangle_points[0][1] + rectangle_points[1][1])/2]

    # 残っているリンクを見て選ばれてないものとの距離を計算。重点からの距離が探索範囲内なら追加候補とする
    remove_keys = []
    for key in link_statuses:
        link_status = link_statuses[key]

        target = 'start'

        if link_status['start'] and link_status['end']:
            continue

        elif link_status['start'] or link_status['end']:
            # どちらか選ばれてたら選ばれてない方をターゲットにする
            target = 'end' if link_status['start'] else 'start'
        else:
            distance1 = calculate_distance(links[key]['start'],center_point) 
            distance2 = calculate_distance(links[key]['end'],center_point)
            target = 'start' if distance1 < distance2 else 'end'


        # 同じリンクの端点が入るのはNG
        if len(list(filter(lambda x: x['key'] == key, point_info_list))) != 0:
            continue
        
        distance = calculate_distance(links[key][target],center_point)

        if (distance < SEARCH_DISTANCE):
            # 入れた状態で四角形とってきて今のポイントの反対側が範囲に入らないか確認
            pre_rectangle_points =  get_rectangle_points(point_info_list + [{'key':key,'target':target}])
            for pi in point_info_list:
                if is_point_inside_rectangle(links[pi['key']]['start' if pi['target'] != 'start' else 'end'],pre_rectangle_points):
                    # 後の作成フローで編集するのでメッセージのみ
                    print(f'{key}({target})が確定済になると確定済リンクの反対側がノード内に入ります')
            is_added = True
            link_statuses[key][target] = True

            # リンクの両端点がいずれかのノードに入っていたらリストから削除して計算量削減(記録)
            if link_statuses[key]['start'] == True and link_statuses[key]['end'] == True:
                remove_keys.append(key)

            point_info_list.append({'key':key,'target':target})
            rectangle_points = pre_rectangle_points
            center_point = [(rectangle_points[0][0] + rectangle_points[1][0])/2,(rectangle_points[0][1] + rectangle_points[1][1])/2]


    # リンクの両端点がいずれかのノードに入っていたらリストから削除して計算量削減(削除)
    for key in remove_keys:
        link_statuses.pop(key)

    if is_added:
        point_info_list = select_points(point_info_list)

    return point_info_list


if __name__ == '__main__':
    
    print('Start set_nodes')

    current_path = os.path.dirname(os.path.abspath(__file__))
    Vissim = com.gencache.EnsureDispatch("Vissim.Vissim")

    input_filepath = os.path.join(current_path, INPUT_FILENAME)

    flag_read_additionally = False
    Vissim.LoadNet(input_filepath, flag_read_additionally)
    
    node_count = 0

    all_links = Vissim.Net.Links.GetAll()

    link_statuses = {} # {'key': {'start': False, 'end': False}}
    links = {} # {'key': {'start': [x1,y1],'end': [x2,y2]}}

    # 各リンクの始点と終点を取得してリスト化
    for current_link in all_links:
        link_start = current_link.LinkPolyPts.GetAll()[0]
        link_end = current_link.LinkPolyPts.GetAll()[-1]
        point_start = [link_start.AttValue('X'), link_start.AttValue('Y')]
        point_end = [link_end.AttValue('X'), link_end.AttValue('Y')]

        link_no = current_link.AttValue('No')

        links[link_no] = {
            'start': point_start,
            'end': point_end,
        }

        link_statuses[link_no] = {
            'start': False,
            'end': False
        }

    # 全リンクの端点がノードに追加されるまで
    while len(link_statuses) > 0:
        key = next(iter(link_statuses))
        target_status = link_statuses.pop(key)
        for target in ['start','end']:
            if (target_status[target] != True):
                # 新しく作るノードのベースとなる地点
                base_point_info = {'key':key,'target':target}
                # ノードにいれるポイントを取得
                
                point_info_list = select_points([base_point_info])
                if len(point_info_list) > 0:
                    node_count += 1
                    print(f'set node : {node_count}')
                    rectangle_points = get_rectangle_points(point_info_list,RECT_MARGIN)
                    Vissim.Net.Nodes.AddNode(node_count, 'POLYGON((' + ', '.join((f'{point[0]} {point[1]}' for point in rectangle_points)) + '))' )



    output_filepath = os.path.join(current_path, OUTPUT_FILENAME)
    Vissim.SaveNetAs(output_filepath)

    Vissim = None

    print('Finish set_nodes')

