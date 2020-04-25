using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AI_BAIDU
{
    class SimnetResult
    {
        int id;
        string subject;
        string snida;
        string titlea;
        string snidb;
        string titleb;
        Decimal score;
        bool issubA;
        bool issubB;

        public int Id { get => id; set => id = value; }
        public string Subject { get => subject; set => subject = value; }
        public string SNIDA { get => snida; set => snida = value; }
        public string TitleA { get => titlea; set => titlea = value; }
        public string SNIDB { get => snidb; set => snidb = value; }
        public string TitleB { get => titleb; set => titleb = value; }
        public decimal Score { get => score; set => score = value; }
        /// <summary>
        /// 是否被裁切，true为裁切，false为未裁切，
        /// 短文本相似度,最大512字节（256字符）
        /// </summary>
        public bool IssubA { get => issubA; set => issubA = value; }
        public bool IssubB { get => issubB; set => issubB = value; }

    }
}
