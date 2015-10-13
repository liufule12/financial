# coding=utf-8
__author__ = 'Fule Liu'

import sys
import time
import os
import csv
import operator

import xlrd


warm_list = []


def print_warm_info():
    """Print warm information.
    """
    for e in warm_list:
        print(e)
    print("共%d个警告。" % len(warm_list))


def file_util(path):
    """Find the all .xls and .xlsx files in path.
    """
    files = []
    cur_files = os.listdir(path)

    for e in cur_files:
        temp_e = e.split('.')
        if temp_e[1] != 'xls' and temp_e[1] != 'xlsx':
            continue
        filename = path + '/' + e
        files.append(filename)

    return files


def make_kmer_list(k, alphabet):
    # Base case.
    if k == 1:
        return alphabet

    # Handle k=0 from user.
    if k == 0:
        return []

    # Error case.
    if k < 1:
        sys.stderr.write("Invalid k=%d" % k)
        sys.exit(1)

    # Precompute alphabet length for speed.
    alphabet_length = len(alphabet)

    # Recursive call.
    return_value = [kmer + alphabet[i_letter] for kmer in make_kmer_list(k - 1, alphabet)
                    for i_letter in range(alphabet_length)]

    return return_value


def make_upto_kmer_list(k_values, alphabet):
    # Compute the k-mer for each value of k.
    return_value = []
    for k in k_values:
        return_value.extend(make_kmer_list(k, alphabet))

    return return_value


def get_col_num_list():
    """得到excel列数字。
    """
    alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    kmer_list = range(1, 3)
    col_num = make_upto_kmer_list(kmer_list, alphabet)

    return col_num


def check_col_label(eva_list, eva_label_list, filename):
    """检验detail表中综合评论栏列属性是否合法.

    :param eva_list: 读取表中得到的评价列.
    :param eva_label_list: template表中评价列.
    :param filename: 所处理的文件.
    """
    row_num = 3
    for ind, eva in enumerate(eva_list):
        if eva not in eva_label_list:
            col_num = get_col_num_list()
            print_warm_info()
            print("错误, %s文件中，"
                  "综合评价栏：%s （%d行%s列） 不存在于template.xls文件。"
                  "合法评价为：" % (filename, eva, row_num, col_num[ind + 2]), eva_label_list)
            sys.exit(0)


def check_comp(comp, comp_order, comp_list, filename):
    """检验券商合法性。

    :param comp: 券商名。
    :param comp_order: 券商序号。
    :param comp_list: 所有券商。
    :param filename: 处理文件名。
    :return:
    """
    if comp not in comp_list:
        if comp == '合计' and comp_order == '':
            return 'End'
        print_warm_info()
        print("错误, %s文件中，"
              "序号：%d 券商：%s 不存在于template.xls文件。"
              "合法券商为："
              % (filename, comp_order, comp), comp_list)
        sys.exit(0)

    return True


def read_objective_score(filename, comp_list):
    """Get company objective score and sum from comp_server_score file.

    :param filename: comp_server_score file.
    :param comp_list: the company list in template.xls.
    :return: comp_server_score, dict, comp: score.
    """
    book = xlrd.open_workbook(filename)
    sh = book.sheet_by_index(0)

    comp_obj_score = {}
    comp_obj_sum = 0.0

    for rx in range(sh.nrows):
        line = sh.row(rx)
        temp_comp = line[0].value
        temp_score = line[1].value

        if temp_comp not in comp_list:
            print_warm_info()
            print("错误, %s文件中，"
                  "%s (%d行)不在template.xls中. " % (filename, temp_comp, rx + 1))
            sys.exit(0)

        comp_obj_score[temp_comp] = float(temp_score)
        comp_obj_sum += float(temp_score)

    return comp_obj_score, comp_obj_sum


def read_template(filename):
    """Read the template.xls.
    """
    book = xlrd.open_workbook(filename)
    sh = book.sheet_by_index(0)

    # Initialize the eva_list.
    eva_list = sh.row_values(2)
    del eva_list[0:2]

    # Initialize the comp_eva_score, comp_list.
    comp_list = []
    comp_eva_score = {}
    for rx in range(3, sh.nrows - 1):
        line = sh.row(rx)
        temp_comp = line[1].value
        comp_list.append(temp_comp)

        for eva in eva_list:
            comp_eva_score[(temp_comp, eva)] = 0

    return comp_eva_score, sh.nrows, comp_list, eva_list


def add_info(files, comp_eva_score, rows, comp_list, eva_label_list):
    """Add the data in detailed_information fold.

    :param files: list, All files needed read.
    :param comp_eva_score: dict, get from function read_template.
    :param rows: int, template.xls rows.
    :return: :raise:
    """
    for file_num, filename in enumerate(files):
        print("Process", file_num, filename)

        book = xlrd.open_workbook(filename)
        sh = book.sheet_by_index(0)

        # 检验行数。
        if sh.nrows < rows or sh.nrows > rows:
            warm_info = ("警告" + filename + "文件行数与template.xls文件行数不一致。")
            warm_list.append(warm_info)

        # 得到并检验所有评价列合法性。
        eva_list = sh.row_values(2)
        del eva_list[0:2]
        check_col_label(eva_list, eva_label_list, filename)

        # 处理每一个券商行。
        for rx in range(3, sh.nrows - 1):
            line = sh.row(rx)
            comp_order = line[0].value

            del line[0]

            # 得到并检测该公司合法性。
            temp_comp = line[0].value
            if check_comp(temp_comp, comp_order, comp_list, rx) == 'End':
                break
            del line[0]

            temp_sum = 0.0

            # For every cell in a line.
            for ind, eva in enumerate(eva_list):
                if (temp_comp, eva) not in comp_eva_score:
                    col_num = get_col_num_list()
                    print_warm_info()
                    print("错误, %s文件中，"
                          "券商：%s，综合评价：%s（%d行%s列）"
                          "不在template.xls文件中或其格式与template。xls文件不一致。" %
                          (filename, temp_comp, eva, rx + 1, col_num[ind + 2]))
                    sys.exit(0)

                score = sh.cell(rx, ind + 2).value

                # 计算单元格。
                try:
                    if eva != '合计' and score != '' and score != ' ':
                        comp_eva_score[(temp_comp, eva)] += float(score)
                        temp_sum += float(score)
                except:
                    col_num = get_col_num_list()
                    print_warm_info()
                    print("错误，%s文件中"
                          "券商：%s，综合评价：%s（%d行%s列）"
                          "单元格内容非法。" %
                          (filename, temp_comp, eva, rx + 1, col_num[ind + 2]))
                    sys.exit(0)

            comp_eva_score[(temp_comp, '合计')] += temp_sum

        print("%d %s 完成." % (file_num, filename))

    return comp_eva_score


def make_perspective_table(comp_eva_score, eva_list):
    """生成透视图。

    :param comp_eva_score: dict, (comp, eva): score.
    :param eva_list: evaluation list.
    :return: comp_table, eva_sum_dict.
    """
    eva_sum_dict = {}
    comp_table = {}
    len_eva = len(eva_list)

    # Make the add_sum table.
    for comp_eva, score in comp_eva_score.items():
        temp_comp = comp_eva[0]
        temp_eva = comp_eva[1]

        # Set the score according the template eva order in a company.
        if temp_comp not in comp_table:
            comp_table[temp_comp] = [0] * len_eva
        comp_table[temp_comp][eva_list.index(temp_eva)] = score

        # Calculate the sum in every evaluation.
        if comp_eva[1] not in eva_sum_dict:
            eva_sum_dict[temp_eva] = score
        else:
            eva_sum_dict[temp_eva] += score

    return comp_table, eva_sum_dict


def write_perspective_table(write_path, comp_table, eva_sum_dict, comp_list, eva_list):
    """写透视图。

    :param write_path: 写文件路径。
    :param comp_table: 透视图表。
    :param eva_sum_dict: 评价总数字典。
    :param comp_list: 所有券商。
    :param eva_list: 所有评价。
    """
    with open(write_path, 'w', newline='') as csvfile:
        f_writer = csv.writer(csvfile, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL)

        f_writer.writerow(['券商'] + eva_list)

        for temp_comp in comp_list:
            f_writer.writerow([temp_comp] + comp_table[temp_comp])

        f_writer.writerow(['总计'] + [eva_sum_dict[temp_eva] for temp_eva in eva_list])


def write_work_table(write_path, comp_table, eva_sum_dict, sum_score_sorted, eva_list,
                     comp_eva_score, subjective_std, obj_score, obj_std, obj_sum):
    """写工作表。

    :param write_path: 写文件路径。
    :param comp_table: 透视图表。
    :param eva_sum_dict: 评价总数字典。
    :param sum_score_sorted: 排序后的券商得分。
    :param eva_list: 所有评价。
    :param comp_eva_score: （券商，评价）：得分。
    :param subjective_std: 主观标准。
    :param obj_score: 主观合计。
    :param obj_std: 客观标准。
    :param obj_sum: 客观合计。
    """
    with open(write_path, 'w', newline='') as csvfile:
        f_writer = csv.writer(csvfile, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL)

        f_writer.writerow(['券商排名', '券商名称', '总分', '主观合计', '主观标准', '客观合计', '客观标准'] + eva_list[:-1])

        for ind, (temp_comp, temp_score) in enumerate(sum_score_sorted):
            f_writer.writerow([ind + 1, temp_comp, temp_score, comp_eva_score[(temp_comp, '合计')],
                               subjective_std[temp_comp], obj_score[temp_comp], obj_std[temp_comp]] +
                              comp_table[temp_comp][:-1])

        f_writer.writerow(['', '总计', '', eva_sum_dict['合计'], 1000.00, obj_sum, 1000.00] +
                          [eva_sum_dict[temp_eva] for temp_eva in eva_list[:-1]])


def write_obj_rank(write_path, obj_rank_sorted):
    """写表3：客观评价得分排名。
    """
    with open(write_path, 'w', newline='') as csvfile:
        f_writer = csv.writer(csvfile, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL)

        for temp_comp, temp_score in obj_rank_sorted:
            f_writer.writerow([temp_comp, temp_score])


def write_comp_rank(write_path, sum_score_sorted, comp_eva_score,
                    subjective_std, sub_rank_sorted,
                    obj_score, obj_std, obj_rank_sorted):
    """写表2：券商排名信息。
    """
    with open(write_path, 'w', newline='') as csvfile:
        f_writer = csv.writer(csvfile, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL)

        f_writer.writerow(['券商排名', '券商名称', '总分', '占比',
                           '主观合计', '主观标准', '主观分排名',
                           '客观合计', '客观标准', '客观分排名'])
        for ind, (temp_comp, temp_score) in enumerate(sum_score_sorted):
            f_writer.writerow([ind + 1, temp_comp, temp_score, str(temp_score / 10) + '%',
                               comp_eva_score[(temp_comp, '合计')], subjective_std[temp_comp], sub_rank_sorted.index((temp_comp, subjective_std[temp_comp])) + 1,
                               obj_score[temp_comp], obj_std[temp_comp], obj_rank_sorted.index((temp_comp, obj_score[temp_comp])) + 1])


def write_plan(write_path, sum_score_sorted, subjective_std, obj_std):
    """写表1：佣金交易量前13计划文件（不包含国信证券）。
    """

    # 求得前13总分合计，不包含国信证券。
    temp_score_sum = 0.0
    top_n = 13
    for temp_comp, temp_score in sum_score_sorted[:top_n]:
        if temp_comp != '国信证券':
            temp_score_sum += temp_score
        else:
            temp_score_sum += sum_score_sorted[top_n][1]

    # 写前13计划，不包含国信证券。
    temp_rate_sum = 0.0
    with open(write_path, 'w', newline='') as csvfile:
        f_writer = csv.writer(csvfile, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL)

        f_writer.writerow(['2014年3季度佣金交易量计划'])
        f_writer.writerow(['券商排名', '券商名称', '总分合计', '主观标准分', '客观标准分', '27家占比', '交易量计划'])

        under_guoxin = False
        for ind, (temp_comp, temp_score) in enumerate(sum_score_sorted[:top_n]):
            if temp_comp == '国信证券':
                under_guoxin = True
                continue

            temp_rate = round(temp_score / 10, 2)
            temp_rate_sum += temp_rate
            if under_guoxin:
                f_writer.writerow([ind, temp_comp, temp_score, subjective_std[temp_comp], obj_std[temp_comp],
                                  str(temp_rate) + '%', str(temp_score / temp_score_sum * 100) + '%'])
            else:
                f_writer.writerow([ind + 1, temp_comp, temp_score, subjective_std[temp_comp], obj_std[temp_comp],
                                  str(temp_rate) + '%', str(temp_score / temp_score_sum * 100) + '%'])

        if under_guoxin:
            temp_comp, temp_score = sum_score_sorted[top_n]
            temp_rate = round(temp_score / 10, 2)
            temp_rate_sum += temp_rate
            f_writer.writerow([top_n, temp_comp, temp_score, subjective_std[temp_comp], obj_std[temp_comp],
                               str(temp_rate) + '%', str(temp_score / temp_score_sum * 100) + '%'])

        f_writer.writerow(['合计', '', temp_score_sum, '', '', str(temp_rate_sum) + '%', '100.00%'])


if __name__ == '__main__':
    if len(sys.argv) != 5 or sys.argv[1] == '-help':
        print("Usage:\n"
              "python finan.py read_fold_path template_path score1_path output_fold_path\n"
              "用法：\n"
              "python finan.py 多张表所在文件夹名称 template文件名 score1文件名 输出文件夹名称路径(所有参数可以使用绝对路径)")
        sys.exit(0)
    elif len(sys.argv) == 5:
        read_path = sys.argv[1]
        template_path = sys.argv[2]
        score1_path = sys.argv[3]
        output_path = sys.argv[4]

        if not os.path.exists(output_path):
            os.mkdir(output_path)

        pers_path = output_path + '/05pers.csv'
        work_path = output_path + '/04work.csv'
        obj_rank_path = output_path + '/03obj_rank.csv'
        rank_path = output_path + '/02rank.csv'
        plan_path = output_path + '/01plan.csv'

    start_time = time.time()
    print("Start!")

    # For test path.
    # read_path = 'detailed_information2'
    # template_path = 'template.xls'
    # score1_path = 'score1.xlsx'
    # pers_path = 'results/05perspective.csv'
    # work_path = 'results/04work.csv'
    # obj_rank_path = 'results/03obj_rank.csv'
    # rank_path = 'results/02rank.csv'
    # plan_path = 'results/01plan.csv'

    files = file_util(read_path)

    template_map, rows, comp_list, eva_list = read_template(template_path)

    comp_eva_score = add_info(files, template_map, rows, comp_list, eva_list)

    comp_table, eva_sum_dict = make_perspective_table(comp_eva_score, eva_list)

    # Calculate subjective items.
    subjective_std = {}
    for temp_comp in comp_list:
        subjective_std[temp_comp] = round(comp_eva_score[(temp_comp, '合计')] / eva_sum_dict['合计'] * 1000 * 0.7, 2)

    # Calculate objective items.
    obj_score, obj_sum = read_objective_score(score1_path, comp_list)
    obj_std = {}
    for temp_comp, temp_score in obj_score.items():
        obj_std[temp_comp] = round(temp_score / obj_sum * 1000 * 0.3, 2)

    # Calculate sum score.
    sum_score = {}
    for temp_comp in comp_list:
        sum_score[temp_comp] = subjective_std[temp_comp] + obj_std[temp_comp]
    sum_score_sorted = sorted(sum_score.items(), key=operator.itemgetter(1), reverse=True)

    write_perspective_table(pers_path, comp_table, eva_sum_dict, comp_list, eva_list)
    write_work_table(work_path, comp_table, eva_sum_dict, sum_score_sorted, eva_list,
                     comp_eva_score, subjective_std, obj_score, obj_std, obj_sum)

    # Calculate the company rank.
    sub_rank_sorted = sorted(subjective_std.items(), key=operator.itemgetter(1), reverse=True)
    obj_rank_sorted = sorted(obj_score.items(), key=operator.itemgetter(1), reverse=True)

    write_obj_rank(obj_rank_path, obj_rank_sorted)
    write_comp_rank(rank_path, sum_score_sorted, comp_eva_score,
                    subjective_std, sub_rank_sorted, obj_score, obj_std, obj_rank_sorted)
    write_plan(plan_path, sum_score_sorted, subjective_std, obj_std)

    print_warm_info()

    print("The execution time: %lf s" % (time.time() - start_time))
    print("Done. :D")