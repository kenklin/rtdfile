/******************************************************************************
*
*  File: RTDFileThread.cpp
*
*  Date: February 5, 2001
*
*  Description:   This file contains the definition of the methods for the 
*  thread that feeds data to the RealTimeData server.  Currently, this thread
*  simply feeds back the current system time.
*
*  Modifications:
******************************************************************************/
#include "windows.h"
#include "windowsx.h"
#include "RTDFileThread.h"

#pragma warning(disable:4100)

IRTDUpdateEvent* pRTDUpdate = NULL;
DWORD WINAPI RTDFileThread( LPVOID pMarshalStream)
{
   CoInitialize( NULL );
   DWORD dwRetVal = 0;
   HRESULT hr = S_OK;

   //Retrieve the RTDUpdate object
   hr = CoGetInterfaceAndReleaseStream( (IStream*)pMarshalStream, 
      IID_IRTDUpdateEvent, (void**)&pRTDUpdate );

   //Set the heartbeat interval to a little more than our timer interval
   if (pRTDUpdate != NULL){
      pRTDUpdate->AddRef();
      hr = pRTDUpdate->put_HeartbeatInterval( 1200 );
   
      //Initiate a timer
      UINT timerID = SetTimer( NULL, 0, 1000, NULL );

      //Spin a message loop so the thread stays alive, and can receive commands 
      //from the parent thread.
      dwRetVal = MessageLoop();
      
      //Kill the timer
      KillTimer( NULL, timerID );

      //Clean up the RTDUpdate object
      pRTDUpdate->Release();
   }

   CoUninitialize();

   //All done...
   return dwRetVal;
}

/******************************************************************************
*  MessageLoop -- This method controls a standard Windows message loop.
*  Parameters: none
*  Returns: the status code from GetMessage.
******************************************************************************/
WPARAM MessageLoop()
{
   MSG msg;
   HRESULT hr;

   //Only WM_QUIT causes GetMessage to return 0.
   while (GetMessage( &msg, NULL, 0, 0)){
      //switch on the message id
      switch( msg.message ){
      case WM_TIMER:
         hr = pRTDUpdate->UpdateNotify();
         break;
      case WM_COMMAND:
         HANDLE_WM_COMMAND(msg.hwnd, msg.wParam, msg.lParam, ThreadOnCommand);
         break;
      default:
         break;
      }
   }
   return msg.wParam;
}

/******************************************************************************
*  ThreadOnCommand -- This method handles the WM_COMMAND messages.
*  Parameters: hWnd -- handle to the window that received the message
*              id -- the command that was received
*              hwndCtl
*              codeNotify
*  Returns: none
******************************************************************************/
void ThreadOnCommand(HWND hWnd, int id, HWND hwndCtl, UINT codeNotify)
{
   HRESULT hr = S_OK;
   //switch on the command
   switch( id ){
   case WM_TERMINATE:
      hr = pRTDUpdate->Disconnect();
   case WM_SILENTTERMINATE:
      PostQuitMessage( hr );
      break;
   default:
      break;
   }
}
