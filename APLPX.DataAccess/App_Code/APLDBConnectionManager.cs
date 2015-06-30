using System;
using System.Data;
using System.Data.SqlClient;

namespace APLPX.DataAccess.App_Code
{
    public class APLDBConnectionManager
    {

         /// <summary>
        /// A managed SQL transaction and connection pair.
        /// </summary>
        /// <remarks>
        /// When a new transaction is created within this class the lifetime of its SQL connection
        /// and the transaction are controlled by the instance.  For convenience, existing transactions may
        /// be wrapped by an instance of this class.  This allows transactions to be shared across commands.
        /// Shared transactions are wrapped, but not managed and Commit()/Rollback() requests are ignored for them.
        /// </remarks>
        public class APLSqlTransaction : IDisposable
        {
            /// <summary>
            /// When true this SQL connection has been closed.
            /// </summary>
            private bool m_Disposed = false;

            /// <summary>
            /// When true member objects are under lifecycle control of this instanace.
            /// </summary>
            private bool m_Managed = true;

            /// <summary>
            /// An open connection to SQL.
            /// </summary>
            private SqlConnection m_Connection = null;

            /// <summary>
            /// The SQL transaction.
            /// </summary>
            private SqlTransaction m_Transaction = null;

            /// <summary>
            /// The transaction associated with this connection.
            /// </summary>
            public SqlTransaction Transaction
            {
                get { return m_Transaction; }
            }

            /// <summary>
            /// Create a new open SQL connection and begin a transaction.
            /// </summary>
            /// <param name="isoLevel"></param>
            /// <param name="connectionStr"></param>
            public APLSqlTransaction(IsolationLevel isoLevel, string connectionStr)
            {
                m_Connection = new SqlConnection(connectionStr);

                // Open the connection to the database
                m_Connection.Open();

                // Generate a transaction on the database and save it in the command object
                m_Transaction = m_Connection.BeginTransaction(isoLevel);
            }

            /// <summary>
            /// Wrap an existing transaction, but don't manage it.
            /// </summary>
            /// <param name="transaction"></param>
            public APLSqlTransaction(SqlTransaction transaction)
            {
                m_Managed = false;
                m_Transaction = transaction;
            }

            /// <summary>
            /// Copy constructor.
            /// </summary>
            /// <param name="t">The instance to copy from.</param>
            private APLSqlTransaction(APLSqlTransaction t)
            {
                m_Managed = false;
                m_Connection = t.m_Connection;
                m_Transaction = t.m_Transaction;
            }

            /// <summary>
            /// Create a copy of a transaction.  The copy is unmanaged and will not close the connection or commit
            /// the transaction.
            /// </summary>
            /// <returns>A copy of this instance.</returns>
            public APLSqlTransaction Copy()
            {
                return new APLSqlTransaction(this);
            }

            /// <summary>
            /// If transaction is Managed, commit it.
            /// </summary>
            public void Commit()
            {
                if (m_Managed)
                {
                    m_Transaction.Commit();
                }
            }

            /// <summary>
            /// If transaction is Managed, roll it back.
            /// </summary>
            public void Rollback()
            {
                if (m_Managed)
                {
                    m_Transaction.Rollback();
                }
            }

            #region Dispose Pattern

            /// <summary>
            /// dispose the object based on managed resource
            /// </summary>
            /// <param name="disposing">When true dispose of managed resources.</param>
            protected virtual void Dispose(bool disposing)
            {
                try
                {
                    if (!m_Disposed)
                    {
                        if (disposing)
                        {
                            if (m_Managed)
                            {
                                if (m_Transaction != null)
                                {
                                    m_Transaction.Dispose();
                                }

                                if (m_Connection != null)
                                {
                                    m_Connection.Dispose();
                                }
                            }

                            m_Transaction = null;
                            m_Connection = null;
                        }

                        m_Disposed = true;
                    }
                }
                catch
                {
                }
            }

            /// <summary>
            /// Close and free managed resources.
            /// </summary>
            public void Dispose()
            {
                Dispose(true);

                // This object will be cleaned up by the Dispose method. 
                // Therefore, you should call GC.SupressFinalize to 
                // take this object off the finalization queue 
                // and prevent finalization code for this object 
                // from executing a second time.
                GC.SuppressFinalize(this);
            }

            #endregion
        }

        /// <summary>
        /// </summary>
        /// 
        /// <param name="transaction">
        /// </param>
        /// 
        /// <returns>
        /// SqlCommand object upon success; null if failure.
        /// </returns>
        /// 
        static internal SqlCommand GetCommand(SqlTransaction transaction)
        {
            // Create a command object that contains the connection to the database
            SqlCommand command = transaction.Connection.CreateCommand();

            // Store the transaction back into the command
            command.Transaction = transaction;

            //Get default timeout from config
            command.CommandTimeout = Properties.Settings.Default.APLDefaultSQLCommandTimeout;

            return command;
        }

        /// <summary>
        /// </summary>
        /// 
        /// <param name="isoLevel">
        /// </param>
        /// 
        /// <returns>
        /// SqlTransaction object upon success; null if failure.
        /// </returns>
        ///
        static internal APLSqlTransaction GetTransaction(IsolationLevel isoLevel, string connectionStr)
        {
            // Return a transaction object that has an open connection to the database
            return new APLSqlTransaction(isoLevel, connectionStr);
        }           
    }
}
