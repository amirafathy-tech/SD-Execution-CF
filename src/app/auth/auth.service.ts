import { Injectable } from '@angular/core';
import { HttpClient, HttpErrorResponse, HttpHeaders } from '@angular/common/http';
import { Router } from '@angular/router';
import { catchError, tap } from 'rxjs/operators';
import { throwError, BehaviorSubject, of } from 'rxjs';
import { AuthUser } from './auth-user.model';
import { environment } from '../../environments/environment';

export interface AuthResponseBackend {
  access_token: string;
  refresh_token: string;
  id_token: string;
  token_type: string;
  expires_in: number;
}

@Injectable({ providedIn: 'root' })
export class AuthService {

  private clientID = environment.clientID
  private clientSecret = environment.clientSecret

  private authUrl = "https://proxy-server.cfapps.eu10-004.hana.ondemand.com/auth"
  private registerUrl = "https://proxy-server.cfapps.eu10-004.hana.ondemand.com/api/iasusers"
   //private registerUrl = "https://sd.cfapps.us10-001.hana.ondemand.com/iasusers"

  loggedInUser = new BehaviorSubject<AuthUser | null>(null);
  private tokenExpirationTimer: any;

  constructor(private http: HttpClient, private router: Router) { }

  signUp(value: string, familyName: string, givenName: string, userName: string) {
    const headers = new HttpHeaders({
     // 'Authorization': 'Bearer eyJraWQiOiI4SmNWTXduY0dmSC1TcFV1Ti1wcU5nOHpSMVUiLCJhbGciOiJSUzI1NiJ9.eyJzdWIiOiI3ZGU3NjE2YS00MzBiLTQ0NjctODU4Yi05ZjI0NGNlYmU4ZTUiLCJhcHBfdGlkIjoiMGJhNzJlMTYtMTRmNC00ZjhmLWEzNzMtZjZiMzUzMjNiODgzIiwiaXNzIjoiaHR0cHM6Ly9handndnFtNHEudHJpYWwtYWNjb3VudHMub25kZW1hbmQuY29tIiwiZ3JvdXBzIjpbIlJlYWQiLCJVc2VyIl0sImdpdmVuX25hbWUiOiJIYWdhciIsImF1ZCI6ImUwMThmNmYzLWVhZTUtNGJkZC1hMjAwLTkxOTgwYzBlN2I0NyIsInNjaW1faWQiOiI3ZGU3NjE2YS00MzBiLTQ0NjctODU4Yi05ZjI0NGNlYmU4ZTUiLCJ1c2VyX3V1aWQiOiI3ZGU3NjE2YS00MzBiLTQ0NjctODU4Yi05ZjI0NGNlYmU4ZTUiLCJ6b25lX3V1aWQiOiIwYmE3MmUxNi0xNGY0LTRmOGYtYTM3My1mNmIzNTMyM2I4ODMiLCJleHAiOjE3MjkxNzM1ODYsImlhdCI6MTcyOTE2OTk4NiwiZmFtaWx5X25hbWUiOiJOYWJpbCIsImp0aSI6IjZjMGVhZjA5LTA5ZWYtNDFlNC1hOWI0LWRiNmRhZGFjNWYxNCIsImVtYWlsIjoiaGFnYXJuYWJpbDdAZ21haWwuY29tIn0.5hUPS9uW8TCj-JWYYoSoLIzRxg-GT1_U3WxOBxl7jRHXuLrOKibwAtqdLdD2cHRFtGS4FFjoRNAs0BshOnkpWI91EPHfQu8KK82HUORWMIlt3rmQ59PPHoIviL3jmYiqedbCJKxRazCQ999tv3xoJdkojkdMNYTmQWLkANUpPtsrXRXUS_9rqK8Am-3WPpFN1a75gqmf9_A_0jwEuGbrKWDtBWM1gLFHyGReGo0IuqmxRK-RI3pUd1qyht5wYA2qj7QrVYnb5jCY0ilLsVLqqJdf8-0Y9qmtw6fB2ivcPGAb8blguWkovxZrnTxHAK7cN50hrV41v98GPxQavC8GpUPXvzxcfYWO1nXhBmk7gVdzmEL6YvWFsWqtylL7hdGyoas26MU1tslUH6OVaUXlymCaTUDKoak3LgH-txiaD54_VZy06VEhYARNRcVjsyMwznA2VzmYXDl1cLYtnW4CfFJEkFeLuUYzvIb4pEaJil9ea2FROV0NeTjHp3WPwHuNnUQdKbX3WBqLkWlqtiU1DkWCvvY9tEmOL9rOJeNBvaWnyyg4YYy2jJH52A2tDMckBCZvG-_QjsEHMZ3WMDUmPU8-hr0ye9BckqmUnraG9LYZzLelTaZzNTeanZKRpIfTehEayYqwocGmuCUD1nSOYAaQLIgGfsmRIdQEoWSw90Y' ,
      'Content-Type': 'application/json'
    });

    const data = {
      'value': value,
      'familyName': familyName,
      'givenName': givenName,
      'userName': userName
    }
    const config = {
      maxBodyLength: Infinity,
      headers,
      body: JSON.stringify(data)
    };
    console.log(data)
   return this.http.post<any>( this.registerUrl,data, { headers })
      // .pipe(
      //   catchError(error => {
      //     console.error(error);
      //     return of(null);
      //   })
      // )
      // .subscribe(response => {
      //   if (response) {
      //     console.log(response);
      //     if (response) {
      //       console.log(response);
      //       this.router.navigate(['/servicetype']);

      //     } else if (response.error_description === "User authentication failed.") {
      //       console.log("40000000");
      //     } else {
      //       console.log(response);
      //     }
      //   }

      // });
  }
  // will be integrated with SAP Auth idenetity service
  signIn(email: string, password: string) {

    const headers = new HttpHeaders({
      'Authorization': 'Basic ' + btoa(`${this.clientID}:${this.clientSecret}`),
      'Content-Type': 'application/x-www-form-urlencoded',
    });

    console.log(headers)
    const data = new URLSearchParams();
    data.set('grant_type', 'password');
    data.set('username', email);
    data.set('password', password);

    console.log(data)
    console.log(data.toString())

    return this.http.post<AuthResponseBackend>(this.authUrl,data.toString(), { headers })
    

    // .pipe(tap(resData => {
    //     console.log(resData.id_token)
    //     const user = new AuthUser(email, resData.id_token);
    //     localStorage.setItem('token', resData.id_token);
    //     this.loggedInUser.next(user);
    //     return user;
    //   }));
  }

  logout() {
    this.loggedInUser.next(null);
    this.router.navigate(['/login']);
    localStorage.removeItem('token');
    if (this.tokenExpirationTimer) {
      clearTimeout(this.tokenExpirationTimer);
    }
    this.tokenExpirationTimer = null;
  }

  autoLogout(expirationDuration: number) {
    this.tokenExpirationTimer = setTimeout(() => {
      this.logout();
    }, expirationDuration);
  }

  private handleError(errorRes: HttpErrorResponse) {
    let errorMessage = 'An unknown error occurred!';
    if (!errorRes.error || !errorRes.error.error) {
      return throwError(errorMessage);
    }
    switch (errorRes.error.error.message) {
      case 'EMAIL_EXISTS':
        errorMessage = 'This email exists already';
        break;
      case 'EMAIL_NOT_FOUND':
        errorMessage = 'This email does not exist.';
        break;
      case 'INVALID_PASSWORD':
        errorMessage = 'This password is not correct.';
        break;
    }
    return throwError(errorMessage);
  }

}