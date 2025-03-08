from __future__ import print_function
import os
import win32com.client as com
import numpy as np
import math

# 設定項目
# ここを必要に応じて書き換えてください
## 入力ファイル名
INPUT_FILENAME = 'input.inpx'
## 出力ファイル名
OUTPUT_FILENAME = 'output.inpx'


# 以下、プログラム

# 2点間の角度
def calculate_angle(x1, y1, x2, y2):
    delta_x = x2 - x1
    delta_y = y2 - y1
    angle_rad = math.atan2(delta_y, delta_x)

    angle_deg = math.degrees(angle_rad)

    return angle_deg

# 回転行列を計算
def rotation_matrix(theta):

    theta_rad = np.radians(theta)

    cos_theta = np.cos(theta_rad)
    sin_theta = np.sin(theta_rad)

    rotation_matrix = np.array([[cos_theta, -sin_theta],
                               [sin_theta, cos_theta]])

    return rotation_matrix


# ノードごとのリンクのペアの生成
def calculate_edge_relation(all_edges):

    for cur_edge in all_edges:
        if cur_edge.AttValue('IsTurn'): # 交差点のEdgeを除外
            continue
        from_node = str(cur_edge.FromNode.AttValue('No'))
        exit_link[from_node].append(cur_edge.LinkSeq.GetAll()[0])
        to_node = str(cur_edge.ToNode.AttValue('No'))
        enter_link[to_node].append(cur_edge.LinkSeq.GetAll()[-1])
        for cur_link in cur_edge.LinkSeq.GetAll():
            link_no = cur_link.AttValue('No')
            related_node[link_no] = [from_node, to_node]

    node_links = {}
    for cur_node in node_list:
        node_no = str(cur_node.AttValue('No'))
        append_list = []
        
        if len(enter_link[node_no]) >= len(exit_link[node_no]):
            for cur_enter in enter_link[node_no]:
                append_list.append([cur_enter, None])
            cur_exit_link_list = exit_link[node_no][:]
            for loop_no in range(len(append_list)):
                cur_list = append_list[loop_no]
                for cur_exit in exit_link[node_no]:
                    enter_no = cur_list[0].AttValue('No')
                    exit_no = cur_exit.AttValue('No')
                    if related_node[enter_no][0] == related_node[exit_no][1]:
                        append_list[loop_no][1] = cur_exit
                        if (cur_exit in cur_exit_link_list):
                            cur_exit_link_list.remove(cur_exit)
                        break
            for cur_exit in cur_exit_link_list:
                append_list.append([None, cur_exit])

            node_links[node_no] = append_list
        else:
            for cur_exit in exit_link[node_no]:
                append_list.append([None, cur_exit])
            cur_enter_link_list = enter_link[node_no][:]
            for loop_no in range(len(append_list)):
                cur_list = append_list[loop_no]
                for cur_enter in enter_link[node_no]:
                    exit_no = cur_list[1].AttValue('No')
                    enter_no = cur_enter.AttValue('No')
                    if related_node[enter_no][0] == related_node[exit_no][1]:
                        append_list[loop_no][0] = cur_enter
                        if (cur_enter in cur_enter_link_list):
                            cur_enter_link_list.remove(cur_enter)
                        break
            for cur_enter in cur_enter_link_list:
                append_list.append([cur_enter, None])
            node_links[node_no] = append_list
    return node_links

# 入リンクをX軸に移動させて、出リンクの入口座標をそれに合わせて移動後、角度を計算
def calculate_link_angle(from_link, to_link):

    from_enter = from_link.LinkPolyPts.GetAll()[-2]
    from_exit = from_link.LinkPolyPts.GetAll()[-1]
    ori_point = [from_enter.AttValue('X'), from_enter.AttValue('Y')]
    des_point = [from_exit.AttValue('X'), from_exit.AttValue('Y')]


    to_enter = to_link.LinkPolyPts.GetAll()[0]
    target_point = [to_enter.AttValue(
        'X')-from_exit.AttValue('X'), to_enter.AttValue('Y')-from_exit.AttValue('Y')]

    angle = calculate_angle(
        des_point[0], des_point[1], ori_point[0], ori_point[1])

    theta = 180-angle
    original_coordinates = np.array(target_point)
    rotation_matrix_30_deg = rotation_matrix(theta)
    rotated_coordinates = np.dot(
        rotation_matrix_30_deg, original_coordinates)

    angle_rad = np.arcsin(
        rotated_coordinates[1] / np.sqrt(rotated_coordinates[0]**2 + rotated_coordinates[1]**2))
    angle_deg = math.degrees(angle_rad)

    return angle_deg

# 直線、右左折の判定
def direction_decision(angle):
    if angle < 22.5 and angle > -22.5:
        state = 'straight'
    elif angle >= 22.5 and angle <= 135:
        state = 'left'
    elif angle <= -22.5 and angle >= -135:
        state = 'right'
    else:
        state = 'other'
    return state


# コネクタをつなげるLinkの組み合わせを取得
def connect_link(cur_links, from_index, to_index):
    from_link = None
    to_link = None

    if cur_links[from_index][0] != None:
        from_link = cur_links[from_index][0]
    if cur_links[to_index][1] != None:
        to_link = cur_links[to_index][1]

    return from_link, to_link


# コネクタを生成
def generate_cross(crossing_links):

    links_num = len(crossing_links)

    connector_list = []

    for from_index in range(links_num):
        cur_con_list = []
        for to_index in range(links_num):

            if from_index == to_index:
                continue

            from_link, to_link = connect_link(crossing_links, from_index, to_index)
            if from_link == None:
                break
            if to_link == None:
                continue

            angle = calculate_link_angle(from_link, to_link)
            state = direction_decision(angle)


            from_lane_index = 0
            to_lane_index = 0
            if state == 'left':
                from_lane_index = len(from_link.Lanes.GetAll())-1
                to_lane_index = len(to_link.Lanes.GetAll())-1
                lane_num = 1
            elif state == 'straight':
                lane_num = len(from_link.Lanes.GetAll())
                if lane_num > len(to_link.Lanes.GetAll()):
                    lane_num = len(to_link.Lanes.GetAll())
            elif state == 'right':
                lane_num = 1

            connector = Vissim.Net.Links.AddConnector(0, from_link.Lanes.GetAll()[from_lane_index], from_link.AttValue(
                'Length2D'), to_link.Lanes.GetAll()[to_lane_index], 0, lane_num, 'LINESTRING EMPTY')
            connector.SetAttValue('state', state)
            cur_con_list.append(connector)

        connector_list.append(cur_con_list)

    for cur_con_list in connector_list:
        cur_index = 0
        for cur_con in cur_con_list:
            for lane in cur_con.Lanes.GetAll():
                sh = Vissim.Net.SignalHeads.AddSignalHead(0, lane, 0.1)
                sh.SetAttValue('SignalSwitch', '')
        cur_index += 1


if __name__ == '__main__':

    print('Start set_connect')

    current_path = os.path.dirname(os.path.abspath(__file__))

    Vissim = com.gencache.EnsureDispatch("Vissim.Vissim")

    input_filepath = os.path.join(current_path, INPUT_FILENAME)

    Vissim.LoadNet(input_filepath, False)

    node_list = Vissim.Net.Nodes.GetAll()
    all_links = Vissim.Net.Links.GetAll()

    cur_pos = []
    global enter_link
    enter_link = {}
    global exit_link
    exit_link = {}
    global related_node
    related_node = {}

    for cur_node in node_list:
        node_no = str(cur_node.AttValue('No'))
        enter_link[node_no] = []
        exit_link[node_no] = []

    Vissim.Net.DynamicAssignment.CreateGraph(0)
    all_edges = Vissim.Net.Edges.GetAll()

    # ユーザ定義属性の追加
    try:
        UDA = Vissim.Net.UserDefinedAttributes.AddUserDefinedDataAttribute(
            0, 'Link', 'state', 'state', 4, 0)
        UDA.SetAttValue('DefValue', '')
        UDA = Vissim.Net.UserDefinedAttributes.AddUserDefinedDataAttribute(
            0, 'SignalHead', 'SignalSwitch', 'SignalSwitch', 4, 0)
        UDA.SetAttValue('DefValue', '')
    except:
        pass

    global node_links
    node_links = calculate_edge_relation(all_edges)

    for cur_node in node_list:
        node_no = str(cur_node.AttValue('No'))
        print(f'node no : {node_no}')

        cur_links = node_links[node_no]
        links_num = len(cur_links)

        if links_num == 3: # T字路
            generate_cross(cur_links)
        elif links_num == 4: # 十字路
            generate_cross(cur_links)

    output_filepath = os.path.join(current_path, OUTPUT_FILENAME)
    Vissim.SaveNetAs(output_filepath)

    Vissim = None

    
    print('Finish set_connect')
