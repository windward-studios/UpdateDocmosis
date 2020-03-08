
namespace UpgradeDocmosis
{
	/// <summary>
	/// Used to set code coverage breakpoints in the code in DEBUG mode only.
	/// </summary>
	public static class Trap
	{
		private static bool stopOnBreak;

		static Trap()
		{
#if DEBUG
			stopOnBreak = true;
#endif // DEBUG
		}

		/// <summary>Will break in to the debugger (debug builds only).</summary>
		public static void trap()
		{
#if DEBUG
			if (stopOnBreak)
				System.Diagnostics.Debugger.Break();
#endif
		}

		/// <summary>Will break in to the debugger if breakOn is true (debug builds only).</summary>
		/// <param name="breakOn">Will break if this boolean value is true.</param>
		public static void trap(bool breakOn)
		{
#if DEBUG
			if (stopOnBreak && breakOn)
				System.Diagnostics.Debugger.Break();
#endif
		}
	}
}
