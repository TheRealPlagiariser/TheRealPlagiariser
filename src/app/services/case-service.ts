import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';
import { environment } from 'src/environments/environment';

// To access the API URL
const apiUrl = environment.apiUrl;

@Injectable({
  providedIn: 'root'
})
export class CaseService {

  constructor(private http: HttpClient) { }

  listCases(): Observable<any> {
    return this.http.get(`${apiUrl}/cases/list`);
  }
}