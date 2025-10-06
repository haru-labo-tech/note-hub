using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using DxLibDLL;
using HLTStudio.Commons;
using HLTStudio.Drawings;
using HLTStudio.GameCommons;

namespace HLTStudio
{
	public class UProgram
	{
		public void Run()
		{
			for (; ; )
			{
				if (Inputs.ENTER.IsPound())
					break;

				DD.EachFrame();
			}
		}
	}
}
