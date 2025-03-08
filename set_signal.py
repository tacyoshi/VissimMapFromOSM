from __future__ import print_function
import os
import win32com.client as com

# 設定項目
# ここを必要に応じて書き換えてください
# 入力ファイル名
INPUT_FILENAME = 'input.inpx'
# 出力ファイル名
OUTPUT_FILENAME = 'output.inpx'
# 基準となる信号グループが決まっているノードのリスト
BASE_NODE_NUMBERS = ['1','2','3','4']


#  - 1 -> - 3 ->
# a       b      c
#  <- 2 - <- 4 -

# related_node = {1: [a,b],2:[b,a],3:[b,c],4:[c,b]} -> リンクがどのノードからノードにつながっているか
# node_links = {a: [[2,1]],b:[[1,2],[4,3]],c:[[3,4]] } -> ノードにつながっているリンクをペアでまとめたもの
# node_adjacent = {a:{b:[2,1]},b:{a:[1,2],c:[4,3]},c:{b:[3,4]}} -> ノードがどのノードに対してどのリンクペアでつながっているかまとめたもの (2状態管理ノードのみ)
# pair_link = {1:2,2:1,3:4,4:3} -> リンクペアを探しやすくマッピングしたもの
# node_groups : ノードごとに管理. コネクタとリンクを２グループ(base or sub)に分類して２状態を管理する.
#               baseの１つのコネクタの値をそのノードのシグナルグループの基準値とする
#               グループの分け方: Stateのstraightの値を使ってグループ分けする
#               例) {'2': {'base': '10000', 'base_connectors': ['10000', '10001', '10002', '10009', '10010','10011'], 'base_links': ['177', '140', '178', '139'], 'sub_connectors': ['10003', '10004', '10005', '10006', '10007', '10008'], 'sub_links': ['176', '2', '175', '1']}
# 2状態管理ノード -> node_linksで見た時にノードペアが 3 or 4 のもの　かつ　node_groupsでsub_connectorsが1つ以上ある場合
# queue_logs -> queueに初めて入ったノードの順番を保存.状態の投票が同数の場合は入った順番が早い方(BASE_NODEに近い状態)を優先する.

# 以下、プログラム

# ノードに設定されているグループを取得
def get_group_value(node_groups, node_no, link_no=None):
    node_group = node_groups[node_no]
    link = Vissim.Net.Links.ItemByKey(node_group['base'])
    cur_lane = link.Lanes.GetAll()[0]
    sh = cur_lane.SigHeads.GetAll()[0]

    if sh.AttValue('SignalSwitch') == '':
        return None

    return sh.AttValue('SignalSwitch') == 'True' if link_no == None or link_no in node_group['base_links'] else (not (sh.AttValue('SignalSwitch') == 'True'))

# シグナルグループを新しい値に書換
def set_group_value(node_groups, node_no, value):
    node_group = node_groups[node_no]

    link = Vissim.Net.Links.ItemByKey(node_group['base'])
    cur_lane = link.Lanes.GetAll()[0]
    sh = cur_lane.SigHeads.GetAll()[0]

    for con_no in node_group['base_connectors']:
        con = Vissim.Net.Links.ItemByKey(con_no)
        for cur_lane in con.Lanes.GetAll():
            sh = cur_lane.SigHeads.GetAll()[0]
            sh.SetAttValue('SignalSwitch', str(value))

    for con_no in node_group['sub_connectors']:
        con = Vissim.Net.Links.ItemByKey(con_no)
        for cur_lane in con.Lanes.GetAll():
            sh = cur_lane.SigHeads.GetAll()[0]
            sh.SetAttValue('SignalSwitch', str(not value))

# ノードの基準のグループに属しているかどうか判定
def is_base_group(node_groups, node_no, link_no):
    return link_no in node_groups[node_no]['base_links']

# 2状態ノードかの判定
def is_2state_node(node_adjacent, node_groups, node_no):
    if node_no not in node_adjacent:
        return False
    if (node_no not in node_groups) or (len(node_groups[node_no]['sub_connectors']) == 0):
        return False
    return True


# ノードごとのコネクションとリンクのグループ化
def calculate_node_groups(node_list, node_links, pair_link):

    node_groups = {}
    for cur_node in node_list:
        node_no = str(cur_node.AttValue('No'))
        junction_links = node_links[node_no]

        base_connectors = []
        sub_connectors = []
        base_links = set()
        sub_links = set()
        print(f"node_no : {node_no}")

        for cur_links in junction_links:
            from_link = cur_links[0]
            if from_link == None:
                continue
            from_edge = from_link.DynAssignEdges.GetAll()[0]
            turns = from_edge.ToEdges.GetAll()

            connectors = []
            links = set()

            # コネクタをまとめる＋直進のペア先を見つける
            for cur_turn in turns:
                from_link = cur_turn.LinkSeq.GetAll()[0]  # In Link
                cur_link = cur_turn.LinkSeq.GetAll()[1]  # 内部コネクタ
                to_link = cur_turn.LinkSeq.GetAll()[2]  # Out Link

                links.add(str(from_link.AttValue('No')))
                if str(from_link.AttValue('No')) in pair_link:
                    links.add(
                        str(pair_link[str(from_link.AttValue('No'))].AttValue('No')))

                connectors.append(str(cur_link.AttValue('No')))

                if cur_link.AttValue('state') == 'straight':
                    links.add(str(to_link.AttValue('No')))
                    if str(to_link.AttValue('No')) in pair_link:
                        links.add(
                            str(pair_link[str(to_link.AttValue('No'))].AttValue('No')))

            if len(base_links) == 0 or (not links.isdisjoint(base_links)):
                base_connectors += connectors
                base_links = base_links.union(links)
            else:
                sub_connectors += connectors
                sub_links = sub_links.union(links)

        if len(base_connectors) == 0:
            continue

        node_groups[node_no] = {"base": min(base_connectors), "base_connectors": base_connectors, "base_links": list(
            base_links), "sub_connectors": sub_connectors, "sub_links": list(sub_links)}

    return node_groups

# 対向リンクのマッピング
def calculate_pair_link(node_links):
    pair_link = {}
    for key in node_links:
        for cur_pair in node_links[key]:
            if cur_pair[0] != None and cur_pair[1] != None:
                link1_no = str(cur_pair[0].AttValue('No'))
                link2_no = str(cur_pair[1].AttValue('No'))
                pair_link[link1_no] = cur_pair[1]
                pair_link[link2_no] = cur_pair[0]

    return pair_link

# ２状態隣接関係マッピングの生成
def calculate_node_adjacent(related_node, node_links):
    node_adjacent = {}
    for from_node_no in node_links:
        if len(node_links[from_node_no]) in [3, 4]:

            for link_pair in node_links[from_node_no]:
                if link_pair[0] == None:
                    enter_link_no = None
                else:
                    enter_link_no = str(link_pair[0].AttValue('No'))
                
                if link_pair[1] == None:
                    exit_link_no = None
                else:
                    exit_link_no = str(link_pair[1].AttValue('No'))

                if enter_link_no is not None:
                    to_node_no = related_node[enter_link_no][0]

                    if from_node_no not in node_adjacent:
                        node_adjacent[from_node_no] = {}

                    if to_node_no not in node_adjacent[from_node_no]:
                        node_adjacent[from_node_no][to_node_no] = {
                            enter_link_no}
                    else:
                        node_adjacent[from_node_no][to_node_no].add(
                            enter_link_no)

                if exit_link_no is not None:
                    to_node_no = related_node[exit_link_no][1]

                    if from_node_no not in node_adjacent:
                        node_adjacent[from_node_no] = {}

                    if to_node_no not in node_adjacent[from_node_no]:
                        node_adjacent[from_node_no][to_node_no] = {
                            exit_link_no}
                    else:
                        node_adjacent[from_node_no][to_node_no].add(
                            exit_link_no)
    return node_adjacent


# 隣接状態の生成
def calculate_edge_relation(all_edges, node_list):

    enter_link = {}
    exit_link = {}
    related_node = {}

    for cur_node in node_list:
        node_no = str(cur_node.AttValue('No'))
        enter_link[node_no] = []
        exit_link[node_no] = []

    for cur_edge in all_edges:
        if cur_edge.AttValue('IsTurn'):
            continue
        from_node = str(cur_edge.FromNode.AttValue('No'))
        exit_link[from_node].append(cur_edge.LinkSeq.GetAll()[0])
        to_node = str(cur_edge.ToNode.AttValue('No'))
        enter_link[to_node].append(cur_edge.LinkSeq.GetAll()[-1])
        for cur_link in cur_edge.LinkSeq.GetAll():
            link_no = str(cur_link.AttValue('No'))
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
                    enter_no = str(cur_list[0].AttValue('No'))
                    exit_no = str(cur_exit.AttValue('No'))
                    if related_node[enter_no][0] == related_node[exit_no][1]:
                        append_list[loop_no][1] = cur_exit
                        if (cur_exit in cur_exit_link_list):
                            cur_exit_link_list.remove(cur_exit)
                        break
            for cur_exit in cur_exit_link_list:
                append_list.append([None, cur_exit])
        else:
            for cur_exit in exit_link[node_no]:
                append_list.append([None, cur_exit])
            cur_enter_link_list = enter_link[node_no][:]
            for loop_no in range(len(append_list)):
                cur_list = append_list[loop_no]
                for cur_enter in enter_link[node_no]:
                    exit_no = str(cur_list[1].AttValue('No'))
                    enter_no = str(cur_enter.AttValue('No'))
                    if related_node[enter_no][0] == related_node[exit_no][1]:
                        append_list[loop_no][0] = cur_enter
                        if (cur_enter in cur_enter_link_list):
                            cur_enter_link_list.remove(cur_enter)
                        break
            for cur_enter in cur_enter_link_list:
                append_list.append([cur_enter, None])
        node_links[node_no] = append_list
    return node_links, related_node


# シグナルグループの決定
def assign_signalgroups(node_no, node_groups, node_adjacent, queue, queue_logs):
    print(f'----{node_no}')
    value = get_group_value(node_groups, node_no)
    newValue = get_majority(node_no, node_groups, node_adjacent, queue_logs)
    if value != newValue:
        print(f"     {node_no} : {value} -> {newValue}")
        set_group_value(node_groups, node_no, newValue)
        if value == None:
            for new_node_no in node_adjacent[node_no]:
                if is_2state_node(node_adjacent, node_groups, new_node_no) == False:
                    continue
                queue.append(new_node_no)
                if new_node_no not in queue_logs:
                    queue_logs.append(new_node_no)
        else:
            for new_node_no in node_adjacent[node_no]:
                if is_2state_node(node_adjacent, node_groups, new_node_no) == False:
                    continue
                assign_signalgroups(new_node_no, node_groups,
                                    node_adjacent, queue, queue_logs)

# 隣接ノードから指定したノードのシグナルグループを算出
def get_majority(node_no, node_groups, node_adjacent, queue_logs):
    link_values = {}
    print(node_adjacent[node_no])
    for to_node_no in node_adjacent[node_no]:
        if is_2state_node(node_adjacent, node_groups, to_node_no) == False:
            continue
        print(f"to_node_no : {to_node_no}")
        link_no = list(node_adjacent[node_no][to_node_no])[0]
        v = get_group_value(node_groups, to_node_no, link_no)
        if v is None:
            link_values[to_node_no] = v
        else:
            if is_base_group(node_groups, node_no, link_no):
                link_values[to_node_no] = v
            else:
                link_values[to_node_no] = not v
    print(link_values)
    values = list(link_values.values())
    tcount = values.count(True)
    fcount = values.count(False)

    # 隣接ノードの値で多数決,同じならキューに入った順番が早い方にする
    if tcount == fcount:
        # print('=======================')
        node_no = find_first(queue_logs, list(
            ({k: v for k, v in link_values.items() if v != None}).keys()))
        # print(f"      {node_no} {get_group_value(node_groups,node_no,link_no)}->{link_values[node_no]}")
        print(queue_logs)
        print(f"node_no : {node_no}")
        return link_values[node_no]
    elif tcount > fcount:
        return True
    else:
        return False


# 配列A,配列Bで配列Aに一番最初に出てくる配列Bの値を取得
def find_first(a, b):
    for value in b:
        if value in a:
            return value

# シグナルグループの適用
def set_signal_group(Vissim):
    signal_heads = Vissim.Net.SignalHeads.GetAll()
    for cur_sh in signal_heads:
        cur_link = cur_sh.Lane.Link
        if cur_sh.AttValue('SignalSwitch') == 'True':
            if cur_link.AttValue('state') == 'right':
                cur_sh.SetAttValue('SG', '1-'+str(2))
            else:
                cur_sh.SetAttValue('SG', '1-'+str(1))
        elif cur_sh.AttValue('SignalSwitch') == 'False':
            if cur_link.AttValue('state') == 'right':
                cur_sh.SetAttValue('SG', '1-'+str(4))
            else:
                cur_sh.SetAttValue('SG', '1-'+str(3))

# コンフリクトエリアの挙動設定
def set_conflictarea_status(Vissim, node_groups):
    for conf in Vissim.Net.ConflictAreas.GetAll():
        link1_no = str(conf.Link1.AttValue('No'))
        link2_no = str(conf.Link2.AttValue('No'))
        print(f"link1_no : {link1_no}")
        print(f"link1_no : {link2_no}")

        if len(conf.Link1.DynAssignTurns.GetAll()) == 0 or len(conf.Link2.DynAssignTurns.GetAll()) == 0:
            continue

        node1_no = str(conf.Link1.DynAssignTurns.GetAll()
                       [0].ToNode.AttValue('No'))
        node2_no = str(conf.Link2.DynAssignTurns.GetAll()
                       [0].ToNode.AttValue('No'))

        if node1_no != node2_no:
            continue

        link1 = Vissim.Net.Links.ItemByKey(link1_no)
        link2 = Vissim.Net.Links.ItemByKey(link2_no)

        link1_state = link1.AttValue('state')
        link2_state = link2.AttValue('state')

        is_base1 = link1_no in node_groups[node1_no]['base_connectors']
        is_base2 = link2_no in node_groups[node2_no]['base_connectors']

        if is_base1 != is_base2:
            conf.SetAttValue('Status', 4)  # Default
            continue

        if link1_state == '' or link2_state == '':
            conf.SetAttValue('Status', 4)  # Default
            continue

        if link1.FromLink.AttValue('No') == link2.FromLink.AttValue('No'):
            conf.SetAttValue('Status', 4)  # Default
            continue

        if (link1_state == 'straight' and link2_state == 'right') or (link1_state == 'left' and link2_state == 'right'):
            conf.SetAttValue('Status', 1)  # link1を優先
        elif (link1_state == 'right' and link2_state == 'straight') or (link1_state == 'right' and link2_state == 'left'):
            conf.SetAttValue('Status', 2)  # link2を優先
        else:
            conf.SetAttValue('Status', 4)  # Default


if __name__ == '__main__':

    print('Start set_signal')

    current_path = os.path.dirname(os.path.abspath(__file__))

    Vissim = com.gencache.EnsureDispatch("Vissim.Vissim")

    input_filepath = os.path.join(current_path, INPUT_FILENAME)

    Vissim.LoadNet(input_filepath, False)

    node_list = Vissim.Net.Nodes.GetAll()

    Vissim.Net.DynamicAssignment.CreateGraph(0)

    all_edges = Vissim.Net.Edges.GetAll()

    # 計算に必要なものを生成
    node_links, related_node = calculate_edge_relation(all_edges, node_list)
    pair_link = calculate_pair_link(node_links)
    node_adjacent = calculate_node_adjacent(related_node, node_links)
    node_groups = calculate_node_groups(node_list, node_links, pair_link)

    # print("====================")
    queue = []
    queue_logs = []  # 値決定の際に多数決で同数だった場合に値の優先度として使用する

    for node_no in BASE_NODE_NUMBERS:
        if is_2state_node(node_adjacent,node_groups,node_no):
            queue_logs.append(node_no)

    
    for node_no in BASE_NODE_NUMBERS:
        if is_2state_node(node_adjacent,node_groups,node_no):
            for target_node_no in node_adjacent[node_no]:
                if is_2state_node(node_adjacent, node_groups, target_node_no) == False:
                    continue
                queue.append(target_node_no)
                if target_node_no not in queue_logs:
                    queue_logs.append(target_node_no)

    while len(queue) > 0:
        print('====================== pop')
        node_no = queue.pop(0)
        assign_signalgroups(node_no, node_groups,
                            node_adjacent, queue, queue_logs)

    print('--- set signal group')
    # シグナルグループの適用
    set_signal_group(Vissim)

    

    print('--- set conflict area')
    set_conflictarea_status(Vissim, node_groups)

    output_filepath = os.path.join(current_path, OUTPUT_FILENAME)
    Vissim.SaveNetAs(output_filepath)

    Vissim = None

    print('Finish set_signal')
